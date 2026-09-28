import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.utils import path_utils  # noqa: E402
from tests.fakes import FakePrinter, FakeWin32Print  # noqa: E402

SAMPLES = ROOT / "tests" / "samples"


@pytest.fixture(autouse=True)
def isolated_data_dir(tmp_path, monkeypatch):
    """每个测试使用独立的数据/日志目录，绝不写入真实用户目录"""
    monkeypatch.setenv("BDP_DATA_DIR", str(tmp_path / "appdata"))
    path_utils.reset_data_root_cache()
    yield tmp_path / "appdata"
    path_utils.reset_data_root_cache()


@pytest.fixture
def fake_win32print(monkeypatch):
    """替换 printer_environment 使用的 win32print（不接触真实打印机）"""
    from src.core import printer_environment
    fake = FakeWin32Print([FakePrinter("Office Printer"), FakePrinter("Target Printer"),
                           FakePrinter("Mono Simplex", duplex=False, color=False, landscape=False,
                                       papers=(9,), paper_names=("A4",))],
                          default="Office Printer")
    monkeypatch.setattr(printer_environment, "_win32print", lambda: fake)
    monkeypatch.setattr(printer_environment.time, "sleep", lambda s: None)
    return fake


@pytest.fixture
def no_processes(monkeypatch):
    """进程快照返回空，避免依赖真实进程"""
    from src.backends import process_utils
    monkeypatch.setattr(process_utils, "snapshot", lambda names: {})
    monkeypatch.setattr(process_utils, "pid_exists", lambda pid: False)
    monkeypatch.setattr(process_utils, "kill_pids", lambda pids, reason="": list(pids))
    return process_utils
