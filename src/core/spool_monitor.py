"""
Windows 打印队列（Spooler）监视

用途：
1. 确认外部打印程序（SumatraPDF 等）退出后确实向打印机提交了任务：
   SumatraPDF 在某些图片上会直接退出且返回 0，但实际上没有生成打印任务，
   过去会被误记为“成功”。
2. 等待本文档产生的打印任务离开队列（已发送到打印机）后再提交下一个文档，
   避免多个任务同时堆在打印机上。

检测方式：
- 优先使用打印机变更通知（FindFirstPrinterChangeNotification / PRINTER_CHANGE_ADD_JOB），
  即使任务一闪而过也能捕获，此时“没有检测到任务”的结论可信（reliable=True）
- 通知不可用时退化为轮询 EnumJobs；轮询可能漏掉极短的任务，此时“没有检测到任务”
  只作为提示，不判定失败（reliable=False）
"""
import logging
import threading
import time
from typing import Callable, Dict, List, Optional, Set

logger = logging.getLogger("bdp.spool")

LogFunc = Callable[[str], None]

PRINTER_ACCESS_USE = 0x00000008
PRINTER_CHANGE_ADD_JOB = 0x00000100
WAIT_OBJECT_0 = 0

# JOB_STATUS_*
JOB_STATUS_PAUSED = 0x00000001
JOB_STATUS_ERROR = 0x00000002
JOB_STATUS_DELETING = 0x00000004
JOB_STATUS_OFFLINE = 0x00000020
JOB_STATUS_PAPEROUT = 0x00000040
JOB_STATUS_PRINTED = 0x00000080
JOB_STATUS_DELETED = 0x00000100
JOB_STATUS_BLOCKED_DEVQ = 0x00000200
JOB_STATUS_USER_INTERVENTION = 0x00000400
JOB_STATUS_COMPLETE = 0x00001000

JOB_DONE_MASK = JOB_STATUS_PRINTED | JOB_STATUS_DELETED | JOB_STATUS_COMPLETE | JOB_STATUS_DELETING
JOB_PROBLEM_MASK = (JOB_STATUS_ERROR | JOB_STATUS_OFFLINE | JOB_STATUS_PAPEROUT | JOB_STATUS_PAUSED
                    | JOB_STATUS_BLOCKED_DEVQ | JOB_STATUS_USER_INTERVENTION)

# 外部打印程序退出后，等待任务出现在队列中的宽限时间（秒）
SUBMIT_GRACE = 3.0
# 等待单个文档的打印任务离开队列的最长时间（秒）；超时后记录提示并继续下一个
JOB_DONE_TIMEOUT = 600.0


class Win32SpoolBackend:
    """真实 Windows 实现：win32print.EnumJobs + winspool 变更通知"""

    def __init__(self, printer_name: str):
        import win32print  # type: ignore
        self._wp = win32print
        self.printer_name = printer_name
        self._handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": PRINTER_ACCESS_USE})
        self._notify = None
        self._notify_printer = None
        self.supports_notify = False
        try:
            self._open_notification()
        except Exception as e:  # noqa: BLE001 - 通知不可用时退化为轮询
            logger.info("打印机变更通知不可用，改为轮询打印队列: %s", e)
            self._close_notification()

    def _open_notification(self) -> None:
        import ctypes
        from ctypes import wintypes

        class PRINTER_DEFAULTSW(ctypes.Structure):
            _fields_ = [("pDatatype", wintypes.LPWSTR), ("pDevMode", ctypes.c_void_p),
                        ("DesiredAccess", wintypes.DWORD)]

        winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        winspool.OpenPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.HANDLE),
                                          ctypes.POINTER(PRINTER_DEFAULTSW)]
        winspool.OpenPrinterW.restype = wintypes.BOOL
        winspool.ClosePrinter.argtypes = [wintypes.HANDLE]
        winspool.FindFirstPrinterChangeNotification.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                                                wintypes.DWORD, ctypes.c_void_p]
        winspool.FindFirstPrinterChangeNotification.restype = wintypes.HANDLE
        winspool.FindNextPrinterChangeNotification.argtypes = [wintypes.HANDLE,
                                                               ctypes.POINTER(wintypes.DWORD),
                                                               ctypes.c_void_p, ctypes.c_void_p]
        winspool.FindNextPrinterChangeNotification.restype = wintypes.BOOL
        winspool.FindClosePrinterChangeNotification.argtypes = [wintypes.HANDLE]
        kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        kernel32.WaitForSingleObject.restype = wintypes.DWORD
        self._ctypes = ctypes
        self._wintypes = wintypes
        self._winspool = winspool
        self._kernel32 = kernel32

        handle = wintypes.HANDLE()
        defaults = PRINTER_DEFAULTSW(None, None, PRINTER_ACCESS_USE)
        if not winspool.OpenPrinterW(self.printer_name, ctypes.byref(handle), ctypes.byref(defaults)):
            raise OSError(ctypes.get_last_error(), "OpenPrinterW")
        self._notify_printer = handle
        notify = winspool.FindFirstPrinterChangeNotification(handle, PRINTER_CHANGE_ADD_JOB, 0, None)
        if not notify or notify == ctypes.c_void_p(-1).value:
            raise OSError(ctypes.get_last_error(), "FindFirstPrinterChangeNotification")
        self._notify = notify
        self.supports_notify = True

    def _close_notification(self) -> None:
        if self._notify:
            try:
                self._winspool.FindClosePrinterChangeNotification(self._notify)
            except Exception:  # noqa: BLE001
                pass
            self._notify = None
        if self._notify_printer:
            try:
                self._winspool.ClosePrinter(self._notify_printer)
            except Exception:  # noqa: BLE001
                pass
            self._notify_printer = None
        self.supports_notify = False

    def list_jobs(self) -> Dict[int, int]:
        """当前队列中的任务 {JobId: Status}"""
        jobs = self._wp.EnumJobs(self._handle, 0, 9999, 1)
        return {int(j["JobId"]): int(j.get("Status", 0) or 0) for j in jobs}

    def wait_add(self, timeout: float) -> bool:
        """等待“新增任务”通知；返回是否收到"""
        if not self._notify:
            time.sleep(timeout)
            return False
        result = self._kernel32.WaitForSingleObject(self._notify, int(timeout * 1000))
        if result != WAIT_OBJECT_0:
            return False
        change = self._wintypes.DWORD(0)
        # 必须调用 FindNext 复位通知对象，否则会一直处于有信号状态
        self._winspool.FindNextPrinterChangeNotification(self._notify, self._ctypes.byref(change), None, None)
        return bool(change.value & PRINTER_CHANGE_ADD_JOB)

    def close(self) -> None:
        self._close_notification()
        if self._handle is not None:
            try:
                self._wp.ClosePrinter(self._handle)
            except Exception:  # noqa: BLE001
                pass
            self._handle = None


class SpoolTracker:
    """
    跟踪单个文档打印期间目标打印机上新出现的任务。

        tracker.start()
        mark = tracker.checkpoint()
        ... 提交打印 ...
        tracker.submitted_since(mark)   # True / False（确认没有任务）/ None（无法确认）
        tracker.wait_done(timeout, cancel_event)
        tracker.stop()
    """

    def __init__(self, backend, log: Optional[LogFunc] = None, poll_interval: float = 0.1):
        self._backend = backend
        self._log = log or (lambda m: logger.info(m))
        self._poll = poll_interval
        self._cond = threading.Condition()
        self._baseline: Set[int] = set()
        self._seen: List[int] = []           # 本文档期间出现的任务 ID（按出现顺序）
        self._add_events = 0                 # 收到的“新增任务”通知次数（任务可能已瞬间消失）
        self._statuses: Dict[int, int] = {}
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._failed = False
        self._io_lock = threading.Lock()  # 串行化 EnumJobs 调用（监视线程与调用线程共用打印机句柄）

    @property
    def reliable(self) -> bool:
        """是否能可靠地判断“没有提交任务”"""
        return bool(getattr(self._backend, "supports_notify", False)) and not self._failed

    def start(self) -> "SpoolTracker":
        self._baseline = set(self._list_jobs())
        self._thread = threading.Thread(target=self._run, name="SpoolTracker", daemon=True)
        self._thread.start()
        return self

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                added = self._backend.wait_add(self._poll)
                jobs = self._list_jobs()
            except Exception as e:  # noqa: BLE001 - 监视失败不影响打印
                logger.warning("读取打印队列失败，停止监视: %s", e)
                with self._cond:
                    self._failed = True
                    self._cond.notify_all()
                return
            self._merge(jobs, added)

    def _list_jobs(self) -> Dict[int, int]:
        with self._io_lock:
            return self._backend.list_jobs()

    def _merge(self, jobs: Dict[int, int], added: bool = False) -> None:
        with self._cond:
            if added:
                self._add_events += 1
            for job_id in jobs:
                if job_id not in self._baseline and job_id not in self._seen:
                    self._seen.append(job_id)
            self._statuses = jobs
            self._cond.notify_all()

    def refresh(self) -> None:
        """立即读取一次打印队列（不等监视线程的下一轮）"""
        if self._failed:
            return
        try:
            self._merge(self._list_jobs())
        except Exception as e:  # noqa: BLE001
            logger.debug("读取打印队列失败: %s", e)

    def checkpoint(self) -> tuple:
        with self._cond:
            return len(self._seen), self._add_events

    def _submitted(self, mark: tuple) -> bool:
        return len(self._seen) > mark[0] or self._add_events > mark[1]

    def submitted_since(self, mark: tuple, grace: float = SUBMIT_GRACE) -> Optional[bool]:
        """
        自 mark 以来是否有任务进入队列。

        Returns:
            True: 检测到任务；False: 可靠地确认没有任务；None: 没检测到但无法确认（轮询模式/监视失败）
        """
        self.refresh()
        deadline = time.monotonic() + grace
        with self._cond:
            while not self._submitted(mark) and not self._failed:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    break
                self._cond.wait(remaining)
            if self._submitted(mark):
                return True
            return False if self.reliable else None

    @property
    def job_ids(self) -> List[int]:
        with self._cond:
            return list(self._seen)

    def wait_done(self, timeout: float = JOB_DONE_TIMEOUT,
                  cancel_event: Optional[threading.Event] = None) -> Optional[bool]:
        """
        等待本文档期间出现的所有任务离开队列（或标记为已打印/已发送）。

        Returns:
            True: 已完成；False: 超时或被取消；None: 没有可等待的任务，或监视失败
        """
        deadline = time.monotonic() + timeout
        warned = False
        with self._cond:
            if not self._seen:
                return None
            while True:
                pending = {j: self._statuses[j] for j in self._seen
                           if j in self._statuses and not self._statuses[j] & JOB_DONE_MASK}
                if not pending:
                    return True
                if self._failed:
                    return None
                if cancel_event is not None and cancel_event.is_set():
                    return False
                if not warned and any(status & JOB_PROBLEM_MASK for status in pending.values()):
                    warned = True
                    self._log("⚠️ 打印机报告任务异常（可能缺纸、卡纸、脱机或暂停），等待处理后继续…")
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return False
                self._cond.wait(min(remaining, 0.5))

    def stop(self) -> None:
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=2)
            self._thread = None
        try:
            self._backend.close()
        except Exception:  # noqa: BLE001
            pass


def create_spool_tracker(printer_name: str, log: Optional[LogFunc] = None) -> Optional[SpoolTracker]:
    """为目标打印机创建监视器；非 Windows 或无法打开打印机时返回 None（不影响打印）"""
    if not printer_name:
        return None
    try:
        backend = Win32SpoolBackend(printer_name)
    except ImportError:
        return None
    except Exception as e:  # noqa: BLE001
        logger.warning("无法监视打印机 %s 的打印队列: %s", printer_name, e)
        return None
    return SpoolTracker(backend, log)
