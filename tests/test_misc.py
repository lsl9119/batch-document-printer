import os
import stat
from pathlib import Path

import pytest

from src.core import file_prep
from src.core.document_manager import DocumentManager
from src.core.models import Document, DuplexMode, FileType, Orientation, PrintSettings, ColorMode, ScalingMode
from src.core.page_count_manager import PageCountManager, PageCountStatus
from src.handlers import gdi_print, sumatra
from src.handlers.text_handler import TextDocumentHandler, read_text_file
from tests.conftest import SAMPLES


# ---------------- SumatraPDF 参数 ----------------
def test_sumatra_print_settings_full():
    s = PrintSettings(copies=3, duplex=True, duplex_mode=DuplexMode.DUPLEX_SHORT, orientation=Orientation.PORTRAIT,
                      color_mode=ColorMode.COLOR, paper_size="Letter", scaling=ScalingMode.SHRINK)
    assert sumatra.build_print_settings(s) == ["3x", "duplexshort", "portrait", "color", "shrink", "paper=letter"]


def test_sumatra_auto_orientation_and_unknown_paper_omitted():
    parts = sumatra.build_print_settings(PrintSettings(paper_size="A4 210 x 297 mm"))
    assert parts == ["1x", "simplex", "monochrome", "fit"]


def test_sumatra_command_with_chinese_path(tmp_path):
    f = tmp_path / "中文 路径" / "文件 名.pdf"
    cmd = sumatra.build_command(f, "打印机 A", PrintSettings(), exe=Path("S.exe"))
    assert cmd == ["S.exe", "-print-to", "打印机 A", "-print-settings", "1x,simplex,monochrome,fit,paper=A4",
                   "-silent", str(f)]


def test_sumatra_failure_raises(monkeypatch, tmp_path):
    import subprocess
    from src.core.print_context import PrintJobError
    monkeypatch.setattr(sumatra, "sumatra_available", lambda: True)
    monkeypatch.setattr(sumatra, "run_detached", lambda cmd, timeout: (1, "打印机脱机"))
    with pytest.raises(PrintJobError, match="打印机脱机"):
        sumatra.run_print(SAMPLES / "sample.pdf", "P", PrintSettings())

    def timeout(cmd, timeout):
        raise subprocess.TimeoutExpired(cmd, 1)

    monkeypatch.setattr(sumatra, "run_detached", timeout)
    with pytest.raises(PrintJobError, match="超时"):
        sumatra.run_print(SAMPLES / "sample.pdf", "P", PrintSettings())


def test_run_detached_does_not_wait_for_inherited_pipes(tmp_path):
    """子进程继承句柄后父进程退出：run_detached 仍应立即返回（capture_output 会卡住）"""
    import sys
    import time
    script = tmp_path / "spawn.py"
    script.write_text("import subprocess, sys\n"
                      "subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(5)'])\n"
                      "sys.stderr.write('bye')\n", encoding="utf-8")
    start = time.monotonic()
    code, detail = sumatra.run_detached([sys.executable, str(script)], timeout=20)
    assert code == 0 and detail == "bye"
    assert time.monotonic() - start < 4


def test_run_detached_timeout_kills(tmp_path):
    import subprocess
    import sys
    with pytest.raises(subprocess.TimeoutExpired):
        sumatra.run_detached([sys.executable, "-c", "import time; time.sleep(30)"], timeout=0.5)


# ---------------- 文本排版 ----------------
def test_wrap_and_layout():
    assert gdi_print.wrap_line("abcdefghij", len, 4) == ["abcd", "efgh", "ij"]
    assert gdi_print.wrap_line("hello world foo", len, 12) == ["hello world", "foo"]
    assert gdi_print.wrap_line("中文测试文本", len, 4) == ["中文测试", "文本"]
    pages = gdi_print.layout_text("a\nb\nc\fd", len, 10, 2)
    assert pages == [["a", "b"], ["c"], ["d"]]
    assert gdi_print.layout_text("", len, 10, 5) == [[""]]


def test_fit_rect():
    w, h = gdi_print.fit_rect(1000, 500, 2000, 2000, ScalingMode.FIT, 100, 100)
    assert (w, h) == (2000, 1000)
    w, h = gdi_print.fit_rect(100, 50, 2000, 2000, ScalingMode.SHRINK, 100, 100)
    assert (w, h) == (100, 50)


def test_text_encodings():
    text, enc = read_text_file(SAMPLES / "sample_gbk.txt")
    assert "中文测试" in text and enc == "gb18030"
    text, enc = read_text_file(SAMPLES / "sample.txt")
    assert enc == "utf-8-sig" and "第1行" in text
    assert TextDocumentHandler().count_pages(SAMPLES / "sample.txt") == 2


# ---------------- 文件预处理 ----------------
def test_prepared_file_direct_open(tmp_path):
    f = tmp_path / "普通 文件.docx"
    f.write_bytes(b"data")
    with file_prep.PreparedFile(f) as prepared:
        assert prepared.path == f and prepared.temp_dir is None


def test_prepared_file_copies_readonly_long_path_and_cleans(tmp_path, monkeypatch):
    work = tmp_path / "work"
    work.mkdir()
    monkeypatch.setattr(file_prep, "get_temp_work_dir", lambda: work)
    monkeypatch.setattr(file_prep, "LONG_PATH_THRESHOLD", 60)
    f = tmp_path / ("很长的名字" * 16 + ".xlsx")
    f.write_bytes(b"data")
    os.chmod(f, stat.S_IREAD)
    with file_prep.PreparedFile(f) as prepared:
        assert prepared.reason == "路径过长"
        assert prepared.path.parent.parent == work
        assert len(prepared.path.name) <= file_prep.MAX_TEMP_NAME + 5
        assert prepared.path.read_bytes() == b"data"
        assert os.access(prepared.path, os.W_OK)
    assert list(work.iterdir()) == []
    os.chmod(f, stat.S_IREAD | stat.S_IWRITE)


def test_owner_lock_file_triggers_copy(tmp_path):
    f = tmp_path / "report.docx"
    f.write_bytes(b"x")
    (tmp_path / "~$report.docx").write_bytes(b"lock")
    assert file_prep.copy_reason(f) == "文件正被其他程序打开"


def test_network_path_detection():
    assert file_prep.is_network_path(Path("\\\\server\\share\\a.docx"))


def test_prepared_file_missing_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        with file_prep.PreparedFile(tmp_path / "missing.docx"):
            pass


# ---------------- 文档管理 ----------------
def test_add_folder_filters_and_sorts(tmp_path):
    for name in ["b.docx", "a.pdf", "~$a.docx", "c - 副本.xlsx", "x.exe", "d.wps", "e.et", "f.dps", ".hidden.pdf"]:
        (tmp_path / name).write_bytes(b"x" * 20)
    manager = DocumentManager()
    added = manager.add_folder(tmp_path, recursive=True,
                               enabled_file_types={"word": True, "pdf": True, "excel": True, "ppt": True})
    assert [d.file_name for d in added] == ["a.pdf", "b.docx", "c - 副本.xlsx", "d.wps", "e.et", "f.dps"]
    # 重复添加被忽略
    assert manager.add_folder(tmp_path) == []


def test_add_folder_respects_type_filter(tmp_path):
    (tmp_path / "a.xlsx").write_bytes(b"x" * 20)
    (tmp_path / "b.pdf").write_bytes(b"x" * 20)
    added = DocumentManager().add_folder(tmp_path, enabled_file_types={"pdf": True, "excel": False})
    assert [d.file_type for d in added] == [FileType.PDF]


# ---------------- 页数统计 ----------------
def test_page_count_sequential_without_office():
    docs = [Document(file_path=SAMPLES / n) for n in ("sample.pdf", "sample.tiff", "sample.png", "sample.txt",
                                                      "sample.pptx")]
    summary = PageCountManager().calculate_all_pages(docs)
    pages = {r.document.file_name: r.page_count for r in summary.error_files + summary.skipped_files}
    assert summary.success_count == 5, pages
    assert summary.total_pages == 2 + 2 + 1 + 2 + 3


def test_page_count_office_unavailable_is_reported(monkeypatch):
    from src.backends.office_backend import BackendSelection, OfficeKind
    manager = PageCountManager(backend_detector=lambda pref: BackendSelection(
        messages={k: "检测到 WPS Office，但WPS 文字 COM 接口不可用。ProgID: KWPS.Application" for k in OfficeKind}))
    summary = manager.calculate_all_pages([Document(file_path=SAMPLES / "sample.docx")])
    result = (summary.skipped_files + summary.error_files)[0]
    assert result.status == PageCountStatus.SKIPPED_NO_OFFICE
    assert "KWPS.Application" in result.error_message


# ---------------- 进程快照健壮性 ----------------
def test_snapshot_tolerates_oserror_from_psutil(monkeypatch):
    from src.backends import process_utils

    class Proc:
        def __init__(self, pid, name, error=None):
            self.pid, self._name, self._error = pid, name, error

        def name(self):
            if self._error:
                raise self._error
            return self._name

    class FakePsutil:
        Error = type("Error", (Exception,), {})

        @staticmethod
        def process_iter():
            return iter([Proc(1, "System", OSError(87, "Invalid parameter")), Proc(2, "WPS.EXE"),
                         Proc(3, "et.exe"), Proc(4, "explorer.exe")])

    monkeypatch.setattr(process_utils, "psutil", FakePsutil)
    monkeypatch.setattr(process_utils, "PSUTIL_AVAILABLE", True)
    assert process_utils.snapshot({"wps.exe", "et.exe"}) == {2: "wps.exe", 3: "et.exe"}
