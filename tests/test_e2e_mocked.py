"""
端到端测试（mock 版）：真实的 PrintController + 真实处理器 + 真实 PrinterEnvironment，
只把最底层替换为替身：WPS COM 对象、win32print、SumatraPDF 子进程、GDI、进程表。

验证：添加任务 → 识别格式 → 调用正确 Backend(ProgID) → 调用打印 → 任务状态完成
      → WPS 进程退出 → 默认打印机恢复
注意：这不是真实 WPS 环境测试。真实环境验证见 scripts/test_wps_com.py 与 scripts/e2e_real_print.py。
"""
import shutil
from pathlib import Path

import pytest

from src.backends import process_utils
from src.backends.office_backend import OfficeKind, OfficeSession, OfficeSessionPool, select_backends
from src.backends.ms_office_backend import MsOfficeBackend
from src.backends.wps_backend import WpsBackend
from src.core.models import Document, DuplexMode, Orientation, PrintSettings, PrintStatus
from src.core.print_controller import PrintController
from src.handlers import gdi_print, sumatra
from tests.conftest import SAMPLES
from tests.fakes import FakeOfficeApp
from tests.test_backend_detection import _statuses
from tests.test_office_session import ProcessTable


@pytest.fixture
def world(tmp_path, monkeypatch, fake_win32print):
    table = ProcessTable()
    table.install(monkeypatch)
    apps = []
    names = {"KWPS.Application": "wps.exe", "KET.Application": "et.exe", "KWPP.Application": "wpp.exe"}

    def dispatch(progid):
        pid = table.spawn(names[progid])
        app = FakeOfficeApp(progid, active_printer=fake_win32print.default,
                            active_printer_writable=progid != "KWPP.Application")
        app.default_at_print = []
        original_quit = app.Quit

        def quit_():
            original_quit()
            table.procs.pop(pid, None)

        app.Quit = quit_
        apps.append(app)
        return app

    selection = select_backends([WpsBackend(), MsOfficeBackend()], "auto", _statuses())

    def pool_factory(sel, log):
        return OfficeSessionPool(selection, log=log,
                                 session_factory=lambda b, k: OfficeSession(b, k, log=log, dispatch_factory=dispatch))

    sumatra_calls = []

    def fake_run(cmd, timeout):
        sumatra_calls.append({"cmd": cmd, "default": fake_win32print.default})
        return 0, ""

    monkeypatch.setattr(sumatra, "sumatra_available", lambda: True)
    monkeypatch.setattr(sumatra, "run_detached", fake_run)
    gdi_calls = []
    monkeypatch.setattr(gdi_print, "print_text",
                        lambda printer, title, text, copies, log=None: gdi_calls.append((printer, title, copies)) or 2)
    monkeypatch.setattr(gdi_print, "print_image",
                        lambda printer, path, settings, log=None: gdi_calls.append((printer, path.name)) or 1)

    state_file = tmp_path / "state.json"
    from src.core.printer_environment import PrinterEnvironment

    controller = PrintController(
        backend_detector=lambda pref: selection,
        pool_factory=pool_factory,
        environment_factory=lambda printer, settings, switch, log: PrinterEnvironment(
            printer, settings, switch_default=switch, state_file=state_file, log=log),
        logs_dir=tmp_path / "logs",
        inter_job_delay=0,
    )
    (tmp_path / "logs").mkdir()
    return dict(controller=controller, apps=apps, table=table, sumatra=sumatra_calls, gdi=gdi_calls,
                win32=fake_win32print, state_file=state_file, tmp=tmp_path)


def _copy_samples(tmp_path, names_map):
    docs = []
    for src_name, dest_name in names_map:
        dest = tmp_path / "输入 文件" / dest_name
        dest.parent.mkdir(exist_ok=True)
        shutil.copy(SAMPLES / src_name, dest)
        docs.append(Document(file_path=dest))
    return docs


def test_full_batch_routes_to_correct_backends(world):
    docs = _copy_samples(world["tmp"], [
        ("sample.docx", "报告 一.docx"), ("sample.xlsx", "数据 表.xlsx"), ("sample.pptx", "演示 稿.pptx"),
        ("sample.pdf", "说明 书.pdf"), ("sample.png", "图片 1.png"), ("sample.txt", "文本 1.txt"),
        ("sample.docx", "原生 文字.wps"), ("sample.xlsx", "原生 表格.et"), ("sample.pptx", "原生 演示.dps"),
        ("sample.jpg", "照片.jpg"), ("sample.bmp", "位图.bmp"), ("sample.tiff", "多页.tiff"),
    ])
    settings = PrintSettings(printer_name="Target Printer", copies=2, paper_size="A3", duplex=True,
                             duplex_mode=DuplexMode.DUPLEX_LONG, orientation=Orientation.LANDSCAPE)
    controller = world["controller"]
    controller.set_print_settings(settings)
    controller.add_documents_to_queue(docs)
    summary = controller.start_batch_print().result(timeout=60)

    assert (summary.success, summary.failed) == (12, 0), [d.last_error for d in docs]
    assert all(d.print_status == PrintStatus.COMPLETED for d in docs)

    # 每个组件只启动一次（会话复用），ProgID 正确
    progids = sorted(app.progid for app in world["apps"])
    assert progids == ["KET.Application", "KWPP.Application", "KWPS.Application"]
    by_progid = {app.progid: app for app in world["apps"]}
    printed_writer = [Path(p[0]).name for p in by_progid["KWPS.Application"].printed]
    assert printed_writer == ["报告 一.docx", "原生 文字.wps"]
    assert [Path(p[0]).name for p in by_progid["KET.Application"].printed] == ["数据 表.xlsx", "原生 表格.et"]
    assert [Path(p[0]).name for p in by_progid["KWPP.Application"].printed] == ["演示 稿.pptx", "原生 演示.dps"]

    # WPS 演示的 ActivePrinter 不可写：WPS 实例在默认打印机切换之后启动，因此继承目标打印机
    assert by_progid["KWPP.Application"].ActivePrinter == "Target Printer"

    # PDF/图片走 SumatraPDF，参数完整
    cmds = [c["cmd"] for c in world["sumatra"]]
    assert len(cmds) == 5
    pdf_cmd = next(c for c in cmds if c[-1].endswith(".pdf"))
    assert pdf_cmd[1:3] == ["-print-to", "Target Printer"]
    assert pdf_cmd[4] == "2x,duplexlong,landscape,monochrome,fit,paper=A3"
    # TXT 走 GDI
    assert world["gdi"] == [("Target Printer", "文本 1.txt", 2)]

    # 所有 WPS 进程退出、文档全部关闭
    assert all(app.quit_called for app in world["apps"])
    assert all(not app.open_docs for app in world["apps"])
    assert world["table"].procs == {}

    # 默认打印机恢复、恢复状态文件已删除、每用户参数已清除
    assert world["win32"].default == "Office Printer"
    assert world["win32"].set_default_calls == ["Target Printer", "Office Printer"]
    assert not world["state_file"].exists()
    assert world["win32"].printers["Target Printer"].user_devmode is None


def test_office_failure_mid_batch_restarts_component(world):
    docs = _copy_samples(world["tmp"], [("sample.docx", "a.docx"), ("sample.docx", "b.docx"),
                                        ("sample.pdf", "c.pdf")])
    controller = world["controller"]
    controller.set_print_settings(PrintSettings(printer_name="Target Printer"))
    original_dispatch_count = len(world["apps"])

    from tests.fakes import dead_server_error

    controller.add_documents_to_queue(docs)
    # 让第一个 WPS 实例的打印抛出"RPC 服务器不可用"并使进程消失
    import src.backends.office_ops as ops
    original_print = ops.writer_print
    state = {"n": 0}

    def flaky_print(app, doc, copies, log=None):
        state["n"] += 1
        if state["n"] == 1:
            for pid, name in list(world["table"].procs.items()):
                if name == "wps.exe":
                    del world["table"].procs[pid]
            raise dead_server_error()
        return original_print(app, doc, copies, log)

    ops_writer_print = ops.writer_print
    ops.writer_print = flaky_print
    try:
        summary = controller.start_batch_print().result(timeout=60)
    finally:
        ops.writer_print = ops_writer_print
    assert (summary.success, summary.failed) == (2, 1)
    assert "RPC 服务器不可用" in docs[0].last_error
    assert docs[1].print_status == PrintStatus.COMPLETED
    # 第二个文档使用了新启动的 WPS 文字实例
    writers = [a for a in world["apps"][original_dispatch_count:] if a.progid == "KWPS.Application"]
    assert len(writers) == 2
    assert world["table"].procs == {}
    assert world["win32"].default == "Office Printer"


def test_long_path_is_copied_to_temp_and_cleaned(world, monkeypatch):
    from src.core import file_prep
    temp_root = world["tmp"] / "work"
    temp_root.mkdir()
    monkeypatch.setattr(file_prep, "get_temp_work_dir", lambda: temp_root)
    monkeypatch.setattr(file_prep, "LONG_PATH_THRESHOLD", 40)
    docs = _copy_samples(world["tmp"], [("sample.docx", "非常长的文件名" * 3 + ".docx")])
    controller = world["controller"]
    controller.set_print_settings(PrintSettings(printer_name="Office Printer"))
    controller.add_documents_to_queue(docs)
    summary = controller.start_batch_print().result(timeout=60)
    assert summary.success == 1
    opened = world["apps"][0].open_calls[0]
    opened_path = opened[1].get("FileName") or opened[0][0]
    assert str(temp_root) in str(opened_path)
    assert list(temp_root.iterdir()) == []  # 临时副本已删除
    # 目标即默认打印机：不切换
    assert world["win32"].set_default_calls == []
