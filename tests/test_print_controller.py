"""
打印队列/控制器测试（全部 mock，不接触真实打印机）
"""
import threading
import time
from contextlib import contextmanager
from pathlib import Path

import pytest

from src.backends.office_backend import BackendSelection, OfficeKind
from src.backends.wps_backend import WpsBackend
from src.core.models import Document, FileType, PrintSettings, PrintStatus
from src.core.print_context import PrintJobError
from src.core.print_controller import PrintController
from src.core.printer_environment import PrinterEnvironmentBusy
from src.handlers.handler_registry import HandlerRegistry
from src.handlers.base_handler import BaseDocumentHandler


class FakeHandler(BaseDocumentHandler):
    def __init__(self, file_type, exts, behaviour=None):
        self.file_type = file_type
        self.exts = exts
        self.behaviour = behaviour or {}
        self.printed = []

    def get_supported_file_types(self):
        return {self.file_type}

    def get_supported_extensions(self):
        return set(self.exts)

    def print_document(self, file_path, settings, context=None):
        action = self.behaviour.get(file_path.name)
        context.backend_used = f"Fake-{self.file_type.value}"
        if action == "fail":
            raise PrintJobError("模拟失败：WPS 演示无法打开文件")
        if action == "crash":
            raise ZeroDivisionError("unexpected")
        if action == "false":
            return False
        if isinstance(action, threading.Event):
            action.wait(10)  # 模拟 WPS 卡死，直到看门狗结束进程
            raise PrintJobError("RPC 服务器不可用")
        self.printed.append(file_path.name)
        return True

    def count_pages(self, file_path, context=None):
        return 1


class FakeEnv:
    def __init__(self, busy=False):
        self.entered = 0
        self.exited = 0
        self.busy = busy
        self.switch_default = None
        self.default_switched = False
        self.warnings = ["双面打印：当前打印机驱动不支持该设置，已使用驱动默认值"]

    def __call__(self, printer, settings, switch_default, log):
        self.switch_default = switch_default
        return self

    def __enter__(self):
        if self.busy:
            raise PrinterEnvironmentBusy("另一个实例正在打印")
        self.entered += 1
        self.default_switched = self.switch_default
        return self

    def __exit__(self, *exc):
        self.exited += 1
        return False


class FakePool:
    def __init__(self):
        self.killed = []
        self.closed = False
        self.hang_events = {}

    def kill_kind(self, kind, reason):
        self.killed.append(kind)
        event = self.hang_events.get(kind)
        if event:
            event.set()
        return [1]

    def close_all(self):
        self.closed = True


def _docs(tmp_path, names):
    docs = []
    for name in names:
        path = tmp_path / name
        path.write_bytes(b"x" * 20)
        docs.append(Document(file_path=path))
    return docs


def _controller(tmp_path, behaviour=None, env=None, pool=None, timeouts=None, selection=None):
    registry = HandlerRegistry()
    handlers = {
        FileType.WORD: FakeHandler(FileType.WORD, {".docx", ".doc", ".wps"}, behaviour),
        FileType.EXCEL: FakeHandler(FileType.EXCEL, {".xlsx", ".xls", ".et"}, behaviour),
        FileType.PPT: FakeHandler(FileType.PPT, {".pptx", ".ppt", ".dps"}, behaviour),
        FileType.PDF: FakeHandler(FileType.PDF, {".pdf"}, behaviour),
    }
    for h in handlers.values():
        registry.register_handler(h)
    env = env or FakeEnv()
    pool = pool or FakePool()
    controller = PrintController(
        registry=registry,
        backend_detector=lambda pref: selection if selection is not None else BackendSelection(
            chosen={k: WpsBackend() for k in OfficeKind}),
        environment_factory=env,
        pool_factory=lambda selection, log: pool,
        timeouts=timeouts,
        logs_dir=tmp_path,
        inter_job_delay=0,
    )
    controller.set_print_settings(PrintSettings(printer_name="Target Printer", copies=2))
    return controller, handlers, env, pool


def _run(controller, docs):
    controller.add_documents_to_queue(docs)
    return controller.start_batch_print().result(timeout=30)


def test_single_failure_does_not_stop_queue(tmp_path):
    controller, handlers, env, pool = _controller(tmp_path, {"3.pptx": "fail"})
    docs = _docs(tmp_path, ["1.docx", "2.xlsx", "3.pptx", "4.pdf", "5.docx"])
    summary = _run(controller, docs)
    assert (summary.success, summary.failed) == (4, 1)
    assert summary.text() == "成功 4，失败 1"
    assert [d.print_status for d in docs] == [PrintStatus.COMPLETED, PrintStatus.COMPLETED, PrintStatus.ERROR,
                                               PrintStatus.COMPLETED, PrintStatus.COMPLETED]
    assert "WPS 演示无法打开文件" in docs[2].last_error
    assert handlers[FileType.WORD].printed == ["1.docx", "5.docx"]
    assert env.entered == env.exited == 1 and env.switch_default is True
    assert pool.closed
    assert not controller.is_printing


def test_unexpected_exception_and_false_return_are_isolated(tmp_path):
    controller, *_ = _controller(tmp_path, {"a.docx": "crash", "b.pdf": "false"})
    docs = _docs(tmp_path, ["a.docx", "b.pdf", "c.pdf"])
    summary = _run(controller, docs)
    assert (summary.success, summary.failed) == (1, 2)
    assert "ZeroDivisionError" in docs[0].last_error or "unexpected" in docs[0].last_error


def test_missing_file_fails_only_that_task(tmp_path):
    controller, *_ = _controller(tmp_path)
    docs = _docs(tmp_path, ["gone.docx", "ok.pdf"])
    docs[0].file_path.unlink()
    summary = _run(controller, docs)
    assert (summary.success, summary.failed) == (1, 1)
    assert "文件不存在" in docs[0].last_error


def test_pdf_only_batch_does_not_switch_default_printer(tmp_path):
    controller, _, env, _ = _controller(tmp_path)
    _run(controller, _docs(tmp_path, ["a.pdf"]))
    assert env.switch_default is False


def test_print_log_contains_required_fields(tmp_path):
    controller, *_ = _controller(tmp_path, {"bad.xlsx": "fail"})
    summary = _run(controller, _docs(tmp_path, ["中文 文件.docx", "bad.xlsx"]))
    log_text = Path(summary.log_file).read_text(encoding="utf-8")
    for field in ("批次开始", "文件=", "类型=", "Backend=Fake-word", "打印机=Target Printer", "纸张=A4",
                  "单双面=单面", "方向=自动", "份数=2", "色彩=黑白", "成功", "失败", "错误=模拟失败", "耗时=",
                  "中文 文件.docx", "批次结束 成功 1，失败 1", "驱动不支持"):
        assert field in log_text, field


def test_watchdog_kills_hung_office_and_continues(tmp_path):
    hang = threading.Event()
    pool = FakePool()
    pool.hang_events[OfficeKind.WRITER] = hang
    controller, *_ = _controller(tmp_path, {"hang.docx": hang}, pool=pool,
                                 timeouts={OfficeKind.WRITER: 0.2})
    docs = _docs(tmp_path, ["hang.docx", "next.pdf"])
    start = time.monotonic()
    summary = _run(controller, docs)
    assert time.monotonic() - start < 5
    assert OfficeKind.WRITER in pool.killed
    assert docs[0].print_status == PrintStatus.ERROR and "超时" in docs[0].last_error
    assert docs[1].print_status == PrintStatus.COMPLETED
    assert (summary.success, summary.failed) == (1, 1)


def test_cancel_marks_remaining_as_cancelled(tmp_path):
    gate = threading.Event()
    controller, handlers, *_ = _controller(tmp_path)
    original = handlers[FileType.PDF].print_document

    def slow(file_path, settings, context=None):
        gate.wait(5)
        return original(file_path, settings, context)

    handlers[FileType.PDF].print_document = slow
    docs = _docs(tmp_path, ["1.pdf", "2.pdf", "3.pdf"])
    controller.add_documents_to_queue(docs)
    future = controller.start_batch_print()
    time.sleep(0.1)
    controller.cancel_current_print()
    gate.set()
    summary = future.result(timeout=10)
    assert summary.success == 1 and summary.cancelled == 2
    assert docs[2].print_status == PrintStatus.CANCELLED


def test_environment_busy_fails_batch_gracefully(tmp_path):
    controller, *_ = _controller(tmp_path, env=FakeEnv(busy=True))
    docs = _docs(tmp_path, ["a.docx", "b.pdf"])
    summary = _run(controller, docs)
    assert summary.failed == 2 and "另一个实例" in summary.fatal_error
    assert all(d.print_status == PrintStatus.ERROR for d in docs)
    assert not controller.is_printing


def test_cannot_start_twice(tmp_path):
    gate = threading.Event()
    controller, handlers, *_ = _controller(tmp_path)
    handlers[FileType.PDF].print_document = lambda *a, **k: gate.wait(5)
    controller.add_documents_to_queue(_docs(tmp_path, ["a.pdf"]))
    future = controller.start_batch_print()
    with pytest.raises(RuntimeError):
        controller.start_batch_print()
    gate.set()
    future.result(timeout=5)


def test_callbacks_receive_status_changes(tmp_path):
    controller, *_ = _controller(tmp_path)
    seen = []
    finished = []
    controller.set_document_callback(lambda d: seen.append((d.file_name, d.print_status)))
    controller.set_finished_callback(finished.append)
    _run(controller, _docs(tmp_path, ["a.pdf"]))
    assert ("a.pdf", PrintStatus.PRINTING) in seen and ("a.pdf", PrintStatus.COMPLETED) in seen
    assert finished and finished[0].success == 1


def test_no_default_switch_when_no_office_backend_available(tmp_path):
    controller, _, env, _ = _controller(tmp_path, selection=BackendSelection())
    _run(controller, _docs(tmp_path, ["a.docx", "b.pdf"]))
    assert env.switch_default is False
