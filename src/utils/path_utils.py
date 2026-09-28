"""
路径工具模块

区分两类路径：
- 资源路径（只读）：SumatraPDF、图标等随程序分发的文件。PyInstaller 打包后位于
  ``sys._MEIPASS``，开发环境位于项目根目录。
- 用户数据路径（可写）：配置、日志、临时文件。便携模式下位于程序目录，
  程序目录不可写（例如安装在 Program Files）时回落到 ``%LOCALAPPDATA%``。
"""
import os
import sys
import tempfile
from pathlib import Path
from typing import Optional

from src.version import APP_ID

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
_data_root_cache: Optional[Path] = None


def is_frozen() -> bool:
    """是否运行在 PyInstaller 打包环境中"""
    return bool(getattr(sys, "frozen", False))


def get_resource_root() -> Path:
    """只读资源根目录"""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return _PROJECT_ROOT


def get_resource_path(relative_path: str) -> Path:
    """获取随程序分发的资源文件的绝对路径"""
    return get_resource_root() / relative_path


def get_sumatra_pdf_path() -> Path:
    """SumatraPDF 可执行文件路径"""
    return get_resource_path("external/SumatraPDF/SumatraPDF.exe")


def get_app_icon_path() -> Path:
    """应用图标路径"""
    return get_resource_path("resources/app_icon.ico")


def get_executable_dir() -> Path:
    """可执行文件（或开发环境项目根目录）所在目录"""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return _PROJECT_ROOT


def _is_writable_dir(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except OSError:
        return False


def get_data_root() -> Path:
    """
    可写数据根目录。

    优先级：环境变量 BDP_DATA_DIR > 程序目录（便携模式，可写时）> %LOCALAPPDATA%\\BatchDocumentPrinter-WPS
    """
    global _data_root_cache
    if _data_root_cache is not None:
        return _data_root_cache

    override = os.environ.get("BDP_DATA_DIR")
    candidates = []
    if override:
        candidates.append(Path(override))
    candidates.append(get_executable_dir())
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        candidates.append(Path(local_appdata) / APP_ID)
    candidates.append(Path.home() / f".{APP_ID}")
    candidates.append(Path(tempfile.gettempdir()) / APP_ID)

    for candidate in candidates:
        if _is_writable_dir(candidate / "data"):
            _data_root_cache = candidate
            return candidate

    # 理论上不会走到这里；返回临时目录以保证程序可以运行
    _data_root_cache = Path(tempfile.gettempdir()) / APP_ID
    return _data_root_cache


def reset_data_root_cache() -> None:
    """测试用：清除数据目录缓存"""
    global _data_root_cache
    _data_root_cache = None


def get_config_dir() -> Path:
    path = get_data_root() / "data"
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_logs_dir() -> Path:
    path = get_data_root() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_temp_work_dir() -> Path:
    """打印过程中使用的临时目录（路径短、仅 ASCII，便于 COM 打开）"""
    path = Path(tempfile.gettempdir()) / "bdp_wps_work"
    path.mkdir(parents=True, exist_ok=True)
    return path
