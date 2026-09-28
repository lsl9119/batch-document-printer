import json
import os

import pytest

from src.core import printer_environment as pe
from src.core.models import ColorMode, DuplexMode, Orientation, PrintSettings


def _settings(**kw):
    base = dict(printer_name="Target Printer", paper_size="A3", copies=2, duplex=True,
                duplex_mode=DuplexMode.DUPLEX_SHORT, color_mode=ColorMode.COLOR,
                orientation=Orientation.LANDSCAPE)
    base.update(kw)
    return PrintSettings(**base)


def test_default_printer_switched_and_restored(fake_win32print, tmp_path):
    state = tmp_path / "state.json"
    with pe.PrinterEnvironment("Target Printer", _settings(), switch_default=True, state_file=state) as env:
        assert fake_win32print.default == "Target Printer"
        assert env.default_switched
        saved = json.loads(state.read_text(encoding="utf-8"))
        assert saved["original_default"] == "Office Printer" and saved["default_switched"]
    assert fake_win32print.default == "Office Printer"
    assert not state.exists()
    assert fake_win32print.open_handles == 0


def test_default_printer_restored_when_printing_raises(fake_win32print, tmp_path):
    state = tmp_path / "state.json"
    with pytest.raises(RuntimeError):
        with pe.PrinterEnvironment("Target Printer", _settings(), switch_default=True, state_file=state):
            raise RuntimeError("WPP 打印崩溃")
    assert fake_win32print.default == "Office Printer"
    assert not state.exists()


def test_no_switch_when_target_is_already_default(fake_win32print, tmp_path):
    with pe.PrinterEnvironment("Office Printer", _settings(printer_name="Office Printer"),
                               switch_default=True, state_file=tmp_path / "s.json") as env:
        assert not env.default_switched
    assert fake_win32print.set_default_calls == []


def test_no_switch_for_pdf_only_batches(fake_win32print, tmp_path):
    with pe.PrinterEnvironment("Target Printer", _settings(), switch_default=False,
                               state_file=tmp_path / "s.json"):
        assert fake_win32print.default == "Office Printer"


def test_devmode_applied_per_user_and_restored(fake_win32print, tmp_path):
    printer = fake_win32print.printers["Target Printer"]
    assert printer.user_devmode is None
    with pe.PrinterEnvironment("Target Printer", _settings(), switch_default=False,
                               state_file=tmp_path / "s.json") as env:
        dm = printer.user_devmode
        assert (dm.Duplex, dm.Color, dm.Orientation, dm.PaperSize, dm.Copies) == \
            (pe.DMDUP_HORIZONTAL, pe.DMCOLOR_COLOR, pe.DMORIENT_LANDSCAPE, 8, 1)
        assert env.warnings == []
    assert printer.user_devmode is None  # 恢复为"没有每用户设置"
    # 全局设置从未被修改
    assert printer.global_devmode.Duplex == 1


def test_existing_user_devmode_restored_field_by_field(fake_win32print, tmp_path):
    from tests.fakes import FakeDevMode
    printer = fake_win32print.printers["Target Printer"]
    printer.user_devmode = FakeDevMode(Duplex=2, Color=1, PaperSize=9, Orientation=1)
    with pe.PrinterEnvironment("Target Printer", _settings(), switch_default=False, state_file=tmp_path / "s.json"):
        assert printer.user_devmode.PaperSize == 8
    assert (printer.user_devmode.Duplex, printer.user_devmode.Color, printer.user_devmode.PaperSize) == (2, 1, 9)


def test_unsupported_settings_fall_back_to_driver_defaults(fake_win32print, tmp_path):
    settings = _settings(printer_name="Mono Simplex")
    with pe.PrinterEnvironment("Mono Simplex", settings, switch_default=False, state_file=tmp_path / "s.json") as env:
        text = "\n".join(env.warnings)
        assert "双面打印：当前打印机驱动不支持该设置，已使用驱动默认值" in text
        assert "彩色打印" in text and "横向打印" in text and "纸张 A3" in text
        dm = fake_win32print.printers["Mono Simplex"].user_devmode
        assert dm.PaperSize == 9  # 未修改为不支持的 A3


def test_crash_recovery_restores_default_printer(fake_win32print, tmp_path):
    state = tmp_path / "state.json"
    # 模拟上次运行在切换默认打印机后被强制结束
    fake_win32print.default = "Target Printer"
    state.write_text(json.dumps({"pid": 999999, "original_default": "Office Printer",
                                 "default_switched": True, "devmode": None}), encoding="utf-8")
    actions = pe.recover_pending_state(state)
    assert fake_win32print.default == "Office Printer"
    assert any("Office Printer" in a for a in actions)
    assert not state.exists()


def test_recovery_skipped_when_other_instance_alive(fake_win32print, tmp_path, monkeypatch):
    state = tmp_path / "state.json"
    state.write_text(json.dumps({"pid": 12345, "original_default": "Office Printer",
                                 "default_switched": True}), encoding="utf-8")
    monkeypatch.setattr(pe, "_state_owner_alive", lambda state: True)
    fake_win32print.default = "Target Printer"
    assert pe.recover_pending_state(state) == []
    assert fake_win32print.default == "Target Printer"
    with pytest.raises(pe.PrinterEnvironmentBusy):
        with pe.PrinterEnvironment("Target Printer", _settings(), True, state_file=state):
            pass


def test_state_file_kept_if_restore_fails(fake_win32print, tmp_path):
    state = tmp_path / "state.json"
    env = pe.PrinterEnvironment("Target Printer", _settings(), switch_default=True, state_file=state)
    env.__enter__()
    fake_win32print.fail_set_default = True
    assert env.restore() is False
    assert state.exists()  # 下次启动时重试
    fake_win32print.fail_set_default = False
    pe.recover_pending_state(state)
    assert fake_win32print.default == "Office Printer" and not state.exists()


def test_lock_released_after_environment(fake_win32print, tmp_path):
    for _ in range(3):
        with pe.PrinterEnvironment("Target Printer", _settings(), True, state_file=tmp_path / "s.json"):
            pass
    assert pe._ENV_LOCK.acquire(timeout=0.1)
    pe._ENV_LOCK.release()


def test_capabilities_paper_lookup(fake_win32print):
    caps = pe.get_capabilities("Target Printer")
    assert caps.supports_duplex and caps.supports_color and caps.supports_landscape
    assert caps.paper_code_for("A3") == 8 and caps.paper_code_for("letter") == 1
    assert caps.paper_code_for("Unknown Paper") is None


def test_busy_instance_state_file_is_not_deleted(fake_win32print, tmp_path, monkeypatch):
    state = tmp_path / "state.json"
    other = {"pid": 4242, "original_default": "Office Printer", "default_switched": True}
    state.write_text(json.dumps(other), encoding="utf-8")
    monkeypatch.setattr(pe, "_state_owner_alive", lambda s: True)
    with pytest.raises(pe.PrinterEnvironmentBusy):
        with pe.PrinterEnvironment("Target Printer", _settings(), True, state_file=state):
            pass
    assert json.loads(state.read_text(encoding="utf-8")) == other  # 另一个实例的状态文件原样保留


def test_failed_recovery_is_inherited_not_overwritten(fake_win32print, tmp_path):
    state = tmp_path / "state.json"
    # 上次崩溃：原默认为 Office Printer，当前仍被切换为 Mono Simplex；本次恢复时 SetDefaultPrinter 失败
    fake_win32print.default = "Mono Simplex"
    state.write_text(json.dumps({"pid": 999999, "original_default": "Office Printer",
                                 "default_switched": True}), encoding="utf-8")
    fake_win32print.fail_set_default = True
    env = pe.PrinterEnvironment("Target Printer", _settings(), switch_default=True, state_file=state)
    env.__enter__()
    saved = json.loads(state.read_text(encoding="utf-8"))
    assert saved["original_default"] == "Office Printer"  # 没有被当前的 Mono Simplex 覆盖
    fake_win32print.fail_set_default = False
    assert env.restore()
    assert fake_win32print.default == "Office Printer" and not state.exists()


def test_original_default_printer_deleted_is_not_retried_forever(fake_win32print, tmp_path):
    state = tmp_path / "state.json"
    state.write_text(json.dumps({"pid": 999999, "original_default": "Removed Printer",
                                 "default_switched": True}), encoding="utf-8")
    actions = pe.recover_pending_state(state)
    assert any("已不存在" in a for a in actions)
    assert not state.exists()


def test_pid_reuse_is_not_treated_as_busy(monkeypatch):
    """状态文件中的 PID 被其他进程复用：创建时间不一致 → 不是同一个实例"""
    pytest.importorskip("psutil")
    import os as _os
    monkeypatch.setattr(pe, "_process_create_time", lambda pid: 5000.0)
    assert pe._pid_alive(_os.getppid(), create_time=1.0) is False
    assert pe._pid_alive(_os.getppid(), create_time=5000.5) is True
    assert pe._state_owner_alive({"pid": _os.getppid(), "pid_create_time": 1.0}) is False


def test_legacy_default_printer_mode_restored(fake_win32print, tmp_path):
    fake_win32print.legacy_mode = pe._MISSING
    with pe.PrinterEnvironment("Target Printer", _settings(), True, state_file=tmp_path / "s.json"):
        fake_win32print.legacy_mode = 1  # SetDefaultPrinter 可能关闭“让 Windows 管理默认打印机”
    assert fake_win32print.legacy_writes[-1] == pe._MISSING
