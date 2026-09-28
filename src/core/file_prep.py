"""
Office 文档打印前的文件预处理

以下情况会把文件复制到短路径的临时目录后再交给 WPS/Office 打开：
- 网络共享路径（UNC \\\\server\\share 或映射的网络驱动器）
- 路径过长（> 200 字符，COM 应用对长路径支持不稳定）
- 文件正被其他程序（通常是用户自己打开的 WPS）占用
  —— 直接打开可能附着到用户正在编辑的文档，打印后关闭会影响用户

其余情况直接打开原文件（只读方式），保证文档中的相对链接、文件名域等正常。
临时副本在打印结束后删除（finally），程序启动时清理上次残留。
"""
import logging
import os
import shutil
import stat
import time
import uuid
from pathlib import Path
from typing import Optional

from src.utils.path_utils import get_temp_work_dir

logger = logging.getLogger("bdp.file_prep")

LONG_PATH_THRESHOLD = 200
MAX_TEMP_NAME = 80


def is_network_path(path: Path) -> bool:
    text = str(path)
    if text.startswith("\\\\") or text.startswith("//"):
        return True
    drive = os.path.splitdrive(text)[0]
    if drive and os.name == "nt":
        try:
            import ctypes
            DRIVE_REMOTE = 4
            return ctypes.windll.kernel32.GetDriveTypeW(drive + "\\") == DRIVE_REMOTE
        except Exception:  # noqa: BLE001 - 无法判断时按本地处理
            return False
    return False


def is_file_locked(path: Path) -> bool:
    """文件是否被其他进程打开（以独占方式打开失败即认为被占用）"""
    if os.name != "nt":
        return False
    try:
        import win32file  # type: ignore
        import pywintypes  # type: ignore
    except ImportError:
        try:
            with open(path, "r+b"):
                return False
        except PermissionError:
            return True
        except OSError:
            return False
    try:
        handle = win32file.CreateFile(str(path), win32file.GENERIC_READ, 0, None,
                                      win32file.OPEN_EXISTING, 0, None)
        handle.Close()
        return False
    except pywintypes.error as e:
        # 32 = ERROR_SHARING_VIOLATION, 33 = ERROR_LOCK_VIOLATION
        return e.winerror in (32, 33)


def has_owner_lock_file(path: Path) -> bool:
    """WPS/Word 打开文档时会在同目录生成 ~$ 开头的锁文件"""
    name = path.name
    candidates = {f"~${name}", f"~${name[2:]}" if len(name) > 2 else ""}
    return any(c and (path.parent / c).exists() for c in candidates)


def copy_reason(path: Path) -> Optional[str]:
    if is_network_path(path):
        return "网络路径"
    if len(str(path)) > LONG_PATH_THRESHOLD:
        return "路径过长"
    if has_owner_lock_file(path) or is_file_locked(path):
        return "文件正被其他程序打开"
    return None


def _safe_temp_name(path: Path) -> str:
    stem, suffix = path.stem, path.suffix
    if len(stem) > MAX_TEMP_NAME:
        stem = stem[:MAX_TEMP_NAME]
    return f"{stem}{suffix}"


def _make_writable(path: Path) -> None:
    try:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass


def _rmtree(path: Path, attempts: int = 5) -> None:
    def _onerror(func, p, _exc):
        _make_writable(Path(p))
        try:
            func(p)
        except OSError:
            pass

    for i in range(attempts):
        shutil.rmtree(path, onerror=_onerror)
        if not path.exists():
            return
        time.sleep(0.3 * (i + 1))  # WPS 可能仍短暂持有文件句柄
    logger.warning("临时目录未能删除（将在下次启动时清理）: %s", path)


class PreparedFile:
    """
    with PreparedFile(original) as prepared:
        open(prepared.path)
    """

    def __init__(self, original: Path, force_copy: bool = False, log=None):
        self.original = Path(original)
        self.force_copy = force_copy
        self.path: Path = self.original
        self.temp_dir: Optional[Path] = None
        self.reason: Optional[str] = None
        self._log = log or (lambda m: logger.info(m))

    def __enter__(self) -> "PreparedFile":
        if not self.original.exists():
            raise FileNotFoundError(f"文件不存在或无法访问: {self.original}")
        self.reason = "强制复制" if self.force_copy else copy_reason(self.original)
        if self.reason:
            temp_dir = get_temp_work_dir() / uuid.uuid4().hex[:10]
            temp_dir.mkdir(parents=True, exist_ok=True)
            target = temp_dir / _safe_temp_name(self.original)
            try:
                shutil.copy2(self.original, target)
                _make_writable(target)
                self.temp_dir = temp_dir
                self.path = target
                self._log(f"{self.reason}，已复制到临时文件后打印: {target}")
            except OSError as e:
                _rmtree(temp_dir)
                self._log(f"{self.reason}，但复制临时文件失败（{e}），直接以只读方式打开原文件")
                self.path = self.original
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.temp_dir is not None:
            _rmtree(self.temp_dir)
        return False


def cleanup_stale_temp(max_age_hours: float = 12.0) -> int:
    """清理上次运行残留的临时目录"""
    root = get_temp_work_dir()
    removed = 0
    now = time.time()
    for child in root.iterdir():
        try:
            if child.is_dir() and now - child.stat().st_mtime > max_age_hours * 3600:
                _rmtree(child, attempts=1)
                removed += 1
        except OSError:
            continue
    return removed


