"""
打印队列监视、图片预处理、"SumatraPDF 未提交任务"检测与 GDI 回退、逐个等待打印完成
（全部 mock，不接触真实打印机）
"""
import threading
import time
from pathlib import Path

import pytest
from PIL import Image

from src.core.models import Document, PrintSettings, ScalingMode
from src.core.print_context import PrintContext, PrintJobError
from src.core.print_controller import PrintController
from src.core.spool_monitor import JOB_STATUS_PAPEROUT, JOB_STATUS_PRINTED, SpoolTracker
from src.handlers import gdi_print, sumatra
from src.handlers.handler_registry import HandlerRegistry
from src.handlers.image_handler import ImageDocumentHandler
from src.handlers.image_prep import max_long_side_px, prepared_image
from src.handlers.pdf_handler import PDFDocumentHandler

SAMPLES = Path(__file__).parent / "samples"


class FakeSpoolBackend:
    def __init__(self, notify=True, jobs=None):
        self.supports_notify = notify
        self.jobs = dict(jobs or {})
        self._adds = 0
        self._lock = threading.Lock()
        self.closed = False

    def list_jobs(self):
        with self._lock:
            return dict(self.jobs)

    def wait_add(self, timeout):
        time.sleep(min(timeout, 0.005))
        with self._lock:
            if self._adds and self.supports_notify:
                self._adds -= 1
                return True
        return False

    def add_job(self, job_id, status=0, vanish=False):
        with self._lock:
            self._adds += 1
            if not vanish:
                self.jobs[job_id] = status

    def finish(self, job_id):
        with self._lock:
            self.jobs.pop(job_id, None)

    def close(self):
        self.closed = True


def _tracker(backend):
    return SpoolTracker(backend, log=lambda m: None, poll_interval=0.01).start()


# ---------------------------------------------------------------- SpoolTracker

def test_tracker_detects_new_job_and_ignores_existing():
    backend = FakeSpoolBackend(jobs={1: 0})
    tracker = _tracker(backend)
    mark = tracker.checkpoint()
    assert tracker.submitted_since(mark, grace=0.1) is False
    backend.add_job(7)
    assert tracker.submitted_since(mark, grace=2) is True
    assert tracker.job_ids == [7]
    tracker.stop()
    assert backend.closed


def test_tracker_counts_job_that_vanished_before_enumeration():
    backend = FakeSpoolBackend()
    tracker = _tracker(backend)
    mark = tracker.checkpoint()
    backend.add_job(9, vanish=True)  # 任务瞬间发送完毕，EnumJobs 来不及看到
    assert tracker.submitted_since(mark, grace=2) is True
    assert tracker.wait_done(timeout=1) is None  # 没有需要等待的任务
    tracker.stop()


def test_tracker_without_notifications_cannot_confirm_absence():
    tracker = _tracker(FakeSpoolBackend(notify=False))
    assert tracker.submitted_since(tracker.checkpoint(), grace=0.05) is None
    tracker.stop()


def test_wait_done_until_job_leaves_queue():
    backend = FakeSpoolBackend()
    tracker = _tracker(backend)
    backend.add_job(3)
    assert tracker.submitted_since((0, 0), grace=2)
    threading.Timer(0.2, backend.finish, args=(3,)).start()
    begin = time.monotonic()
    assert tracker.wait_done(timeout=5) is True
    assert time.monotonic() - begin >= 0.15
    tracker.stop()


def test_wait_done_treats_printed_status_as_done_and_times_out_otherwise():
    backend = FakeSpoolBackend()
    tracker = _tracker(backend)
    backend.add_job(4, status=JOB_STATUS_PRINTED)
    assert tracker.submitted_since((0, 0), grace=2)
    assert tracker.wait_done(timeout=1) is True
    backend.add_job(5, status=JOB_STATUS_PAPEROUT)
    assert tracker.submitted_since((1, 1), grace=2)
    assert tracker.wait_done(timeout=0.2) is False
    cancel = threading.Event()
    cancel.set()
    assert tracker.wait_done(timeout=5, cancel_event=cancel) is False
    tracker.stop()


# ---------------------------------------------------------------- 图片预处理

def _make_photo(path, size=(5000, 3000), orientation=None, mode="RGB"):
    img = Image.new(mode, size, "white" if mode != "RGBA" else (255, 0, 0, 0))
    exif = Image.Exif()
    if orientation:
        exif[0x0112] = orientation
    img.save(path, "JPEG" if mode != "RGBA" else "PNG", exif=exif.tobytes())
    return path


def test_large_photo_rotated_and_downscaled(tmp_path):
    photo = _make_photo(tmp_path / "IMG20260323151344.jpg", orientation=6)
    notes = []
    settings = PrintSettings(paper_size="A4")
    with prepared_image(photo, settings, lambda m: None, notes) as path:
        assert path != photo and path.exists() and path.name == "IMG20260323151344.jpg"
        with Image.open(path) as out:
            # EXIF 6 = 需顺时针旋转 90°：横图变竖图，长边缩到 A4 300dpi
            assert out.height == max_long_side_px("A4")
            assert abs(out.width - max_long_side_px("A4") * 3000 / 5000) <= 1
            assert out.getexif().get(0x0112) in (None, 1)
        temp_dir = path.parent
    assert not temp_dir.exists()
    assert "转正" in notes[0] and "缩小 5000x3000" in notes[0]
    assert photo.exists()


def test_small_upright_image_used_as_is(tmp_path):
    photo = _make_photo(tmp_path / "small.jpg", size=(800, 600))
    notes = []
    with prepared_image(photo, PrintSettings(), None, notes) as path:
        assert path == photo
    assert notes == []


def test_noscale_only_rotates(tmp_path):
    photo = _make_photo(tmp_path / "p.jpg", size=(5000, 3000), orientation=6)
    with prepared_image(photo, PrintSettings(scaling=ScalingMode.NOSCALE)) as path:
        with Image.open(path) as out:
            assert out.size == (3000, 5000)


def test_transparent_png_composited_on_white(tmp_path):
    png = _make_photo(tmp_path / "t.png", size=(100, 100), mode="RGBA")
    with prepared_image(png, PrintSettings()) as path:
        with Image.open(path) as out:
            assert out.mode == "RGB" and out.getpixel((5, 5)) == (255, 255, 255)


def test_transparent_png_stays_lossless_png(tmp_path):
    png = _make_photo(tmp_path / "t.png", size=(100, 100), mode="RGBA")
    with prepared_image(png, PrintSettings()) as path:
        assert path.suffix == ".png"


def test_mpo_phone_photo_is_processed(tmp_path):
    frame = Image.new("RGB", (5000, 3000), "white")
    exif = Image.Exif()
    exif[0x0112] = 8
    mpo = tmp_path / "IMG_6794.jpeg"
    frame.save(mpo, "MPO", save_all=True, append_images=[frame.copy()], exif=exif.tobytes())
    with Image.open(mpo) as check:
        assert check.format == "MPO" and check.n_frames == 2
    with prepared_image(mpo, PrintSettings()) as path:
        assert path != mpo and path.suffix == ".jpg"
        with Image.open(path) as out:
            assert out.height == 3508 and out.width < out.height


def test_jpeg_draft_decoding_keeps_full_target_size(tmp_path):
    photo = _make_photo(tmp_path / "huge.jpg", size=(8160, 6120))
    with prepared_image(photo, PrintSettings()) as path:
        with Image.open(path) as out:
            assert out.size == (3508, 2631)


def test_multiframe_tiff_untouched():
    tiff = SAMPLES / "sample.tiff"
    with prepared_image(tiff, PrintSettings()) as path:
        assert path == tiff


def test_paper_long_side():
    assert max_long_side_px("A4") == 3508
    assert max_long_side_px("A3") == 4961
    assert max_long_side_px("未知纸张") == 3508


# ---------------------------------------------------------------- 处理器：未提交任务检测

class StubTracker:
    """submitted_since 依次返回给定结果"""

    def __init__(self, results):
        self.results = list(results)
        self.marks = 0

    def checkpoint(self):
        self.marks += 1
        return (0, 0)

    def submitted_since(self, mark, grace=0):
        return self.results.pop(0)


@pytest.fixture
def fake_sumatra(monkeypatch):
    calls = []
    monkeypatch.setattr(sumatra, "sumatra_available", lambda: True)

    def run_print(path, printer, settings, log=None):
        calls.append(path)
        return "1x,simplex"

    monkeypatch.setattr(sumatra, "run_print", run_print)
    return calls


@pytest.fixture
def fake_gdi(monkeypatch):
    calls = []
    monkeypatch.setattr(gdi_print, "print_image", lambda printer, path, settings, log=None: calls.append(path) or 1)
    return calls


def test_image_falls_back_to_gdi_when_sumatra_submits_nothing(tmp_path, fake_sumatra, fake_gdi):
    photo = _make_photo(tmp_path / "IMG_6794.jpeg", size=(400, 300))
    ctx = PrintContext(spool=StubTracker([False, True]))
    assert ImageDocumentHandler().print_document(photo, PrintSettings(printer_name="P"), ctx)
    assert len(fake_sumatra) == 1 and len(fake_gdi) == 1
    assert ctx.backend_used == "Windows GDI"
    assert any("SumatraPDF 失败后改用 GDI" in n for n in ctx.notes)


def test_image_sumatra_ok_when_job_seen(tmp_path, fake_sumatra, fake_gdi):
    photo = _make_photo(tmp_path / "a.jpg", size=(400, 300))
    ctx = PrintContext(spool=StubTracker([True]))
    ImageDocumentHandler().print_document(photo, PrintSettings(printer_name="P"), ctx)
    assert ctx.backend_used == "SumatraPDF" and fake_gdi == []


def test_image_fails_when_neither_engine_submits(tmp_path, fake_sumatra, fake_gdi):
    photo = _make_photo(tmp_path / "a.jpg", size=(400, 300))
    with pytest.raises(PrintJobError, match="没有出现打印任务"):
        ImageDocumentHandler().print_document(photo, PrintSettings(printer_name="P"),
                                              PrintContext(spool=StubTracker([False, False])))


def test_image_unconfirmed_job_is_noted_not_failed(tmp_path, fake_sumatra, fake_gdi):
    photo = _make_photo(tmp_path / "a.jpg", size=(400, 300))
    ctx = PrintContext(spool=StubTracker([None]))
    ImageDocumentHandler().print_document(photo, PrintSettings(printer_name="P"), ctx)
    assert fake_gdi == [] and "未能确认任务已进入打印队列" in ctx.notes


def test_image_handler_prints_prepared_copy(tmp_path, fake_sumatra, fake_gdi):
    photo = _make_photo(tmp_path / "big.jpg", size=(6000, 4000))
    ImageDocumentHandler().print_document(photo, PrintSettings(printer_name="P"), PrintContext())
    assert fake_sumatra[0] != photo and fake_sumatra[0].name == "big.jpg"


def test_pdf_fails_when_sumatra_submits_nothing(fake_sumatra):
    with pytest.raises(sumatra.NoJobSubmitted):
        PDFDocumentHandler().print_document(SAMPLES / "sample.pdf", PrintSettings(printer_name="P"),
                                            PrintContext(spool=StubTracker([False])))


# ---------------------------------------------------------------- 控制器：逐个等待打印完成

class RecordingHandler(ImageDocumentHandler):
    """每次打印向假打印队列添加一个任务"""

    def __init__(self, backend, events):
        self.backend = backend
        self.events = events
        self.next_id = 100

    def print_document(self, file_path, settings, context=None):
        self.events.append(("submit", file_path.name, set(self.backend.list_jobs())))
        self.next_id += 1
        self.backend.add_job(self.next_id)
        job = self.next_id
        threading.Timer(0.3, lambda: (self.backend.finish(job), self.events.append(("done", job)))).start()
        return True


def test_controller_waits_for_previous_job_to_leave_queue(tmp_path):
    backend = FakeSpoolBackend()
    events = []
    registry = HandlerRegistry()
    registry.register_handler(RecordingHandler(backend, events))
    controller = PrintController(
        registry=registry,
        environment_factory=lambda *a: _NullEnv(),
        logs_dir=tmp_path,
        inter_job_delay=0,
        spool_tracker_factory=lambda printer, log: SpoolTracker(backend, log, poll_interval=0.01),
    )
    controller.set_print_settings(PrintSettings(printer_name="Target Printer"))
    docs = []
    for name in ("1.png", "2.png", "3.png"):
        path = tmp_path / name
        Image.new("RGB", (50, 50)).save(path)
        docs.append(Document(file_path=path))
    controller.add_documents_to_queue(docs)
    summary = controller.start_batch_print().result(timeout=30)
    assert summary.success == 3
    submits = [e for e in events if e[0] == "submit"]
    # 每次提交时，打印队列里不应再有上一个文件的任务
    assert all(queue == set() for _, _, queue in submits)
    assert all(any("已等待打印队列发送完成" in n for n in r.notes) for r in summary.records)


class _NullEnv:
    default_switched = False
    warnings = []

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False
