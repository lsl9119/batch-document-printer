"""
Office 进程跟踪工具

通过"创建 COM 对象前后的进程快照差集"识别本程序启动的 WPS/Office 进程，
只对自己启动的进程执行强制结束，绝不误杀用户手动打开的 WPS。
"""
import logging
import sys
import time
from typing import Dict, Iterable, List, Optional, Set, Tuple

logger = logging.getLogger("bdp.process")

try:
    import psutil  # type: ignore
    PSUTIL_AVAILABLE = True
except ImportError:  # pragma: no cover - 打包环境中始终存在
    psutil = None
    PSUTIL_AVAILABLE = False


def tracking_available() -> bool:
    """能否识别进程（Windows 用 Toolhelp32，其它平台依赖 psutil）"""
    return sys.platform == "win32" or PSUTIL_AVAILABLE


def _terminate_native(pid: int) -> bool:
    """psutil 不可用/失败时，直接调用 TerminateProcess"""
    if sys.platform != "win32":
        return False
    import ctypes
    from ctypes import wintypes
    PROCESS_TERMINATE = 0x0001
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
    if not handle:
        return False
    try:
        return bool(kernel32.TerminateProcess(handle, 1))
    finally:
        kernel32.CloseHandle(handle)


def _toolhelp_processes() -> List[Tuple[int, str]]:
    """
    Windows 原生进程快照（CreateToolhelp32Snapshot）。
    与 psutil.Process.name() 不同，它不需要打开目标进程，对提权进程/受保护进程也能拿到映像名。
    """
    import ctypes
    from ctypes import wintypes

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
                    ("th32ProcessID", wintypes.DWORD), ("th32DefaultHeapID", ctypes.c_size_t),
                    ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
                    ("th32ParentProcessID", wintypes.DWORD), ("pcPriClassBase", ctypes.c_long),
                    ("dwFlags", wintypes.DWORD), ("szExeFile", ctypes.c_wchar * 260)]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    kernel32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    TH32CS_SNAPPROCESS = 0x00000002
    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if not snap or snap == wintypes.HANDLE(-1).value:
        raise OSError(ctypes.get_last_error(), "CreateToolhelp32Snapshot failed")
    result: List[Tuple[int, str]] = []
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = kernel32.Process32FirstW(snap, ctypes.byref(entry))
        while ok:
            result.append((int(entry.th32ProcessID), entry.szExeFile))
            ok = kernel32.Process32NextW(snap, ctypes.byref(entry))
    finally:
        kernel32.CloseHandle(snap)
    return result


def _psutil_processes() -> List[Tuple[int, str]]:
    result: List[Tuple[int, str]] = []
    if not PSUTIL_AVAILABLE:
        return result
    # 不使用 process_iter(attrs)：个别系统进程读取名称时会抛出 psutil 以外的 OSError，
    # 在 attrs 模式下会中断整个迭代
    for proc in psutil.process_iter():
        try:
            result.append((proc.pid, proc.name() or ""))
        except (psutil.Error, OSError):
            continue
    return result


def list_processes() -> List[Tuple[int, str]]:
    """[(pid, 映像名)]：Windows 上优先 Toolhelp32，失败或非 Windows 时用 psutil"""
    if sys.platform == "win32":
        try:
            return _toolhelp_processes()
        except (OSError, AttributeError, ValueError) as e:
            logger.debug("Toolhelp32 进程快照失败，改用 psutil: %s", e)
    return _psutil_processes()


def snapshot(process_names: Iterable[str]) -> Dict[int, str]:
    """返回 {pid: 进程名(小写)}，仅包含指定进程名（不区分大小写）"""
    names = {n.lower() for n in process_names}
    return {pid: name.lower() for pid, name in list_processes() if name.lower() in names}


def pid_exists(pid: int) -> bool:
    if sys.platform == "win32":
        try:
            return any(p == pid for p, _ in _toolhelp_processes())
        except (OSError, AttributeError, ValueError):
            pass
    if not PSUTIL_AVAILABLE:
        return False
    try:
        return psutil.pid_exists(pid)
    except (psutil.Error, OSError):
        return False


def wait_for_exit(pids: Iterable[int], timeout: float) -> Set[int]:
    """等待进程退出，返回超时后仍存活的 pid"""
    remaining = set(pids)
    if not remaining:
        return set()
    deadline = time.monotonic() + timeout
    while remaining and time.monotonic() < deadline:
        remaining = {pid for pid in remaining if pid_exists(pid)}
        if remaining:
            time.sleep(0.2)
    return remaining


def kill_pids(pids: Iterable[int], reason: str = "") -> List[int]:
    """强制结束指定进程（及其子进程），返回实际结束的 pid 列表"""
    killed: List[int] = []
    for pid in set(pids):
        done = False
        if PSUTIL_AVAILABLE:
            try:
                proc = psutil.Process(pid)
                try:
                    children = proc.children(recursive=True)
                except (psutil.Error, OSError):
                    children = []
                for child in children:
                    try:
                        child.kill()
                    except (psutil.Error, OSError):
                        pass
                proc.kill()
                done = True
            except psutil.NoSuchProcess:
                continue
            except (psutil.Error, OSError) as e:
                logger.debug("psutil 结束进程 pid=%s 失败，改用 TerminateProcess: %s", pid, e)
        if not done:
            try:
                done = _terminate_native(pid)
            except OSError as e:
                logger.error("结束进程 pid=%s 失败: %s", pid, e)
        if done:
            killed.append(pid)
            logger.warning("已强制结束进程 pid=%s %s", pid, reason)
    return killed


def list_running(process_names: Iterable[str]) -> List[str]:
    """返回正在运行的指定进程描述列表，用于日志与 GUI 提示"""
    return [f"{name}(pid={pid})" for pid, name in sorted(snapshot(process_names).items())]


def find_new_pids(before: Dict[int, str], after: Optional[Dict[int, str]] = None,
                  process_names: Iterable[str] = ()) -> Set[int]:
    """快照差集"""
    if after is None:
        after = snapshot(process_names)
    return set(after) - set(before)
