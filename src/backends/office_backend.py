"""
Office 引擎抽象层

统一封装 WPS Office（KWPS/KET/KWPP）与 Microsoft Office（Word/Excel/PowerPoint）
的 COM 生命周期，处理器（handlers）只依赖本模块，不直接硬编码 ProgID。

核心对象：
- OfficeKind        组件类型：文字 / 表格 / 演示
- OfficeBackend     某个 Office 套件（WPS 或 MS Office）：ProgID、进程名、检测、创建会话
- OfficeSession     一个 COM Application 实例：启动、配置、退出、强制结束
- OfficeSessionPool 批量任务期间复用的会话池（同一线程内使用）
- Watchdog          超时看门狗：超时后强制结束本程序启动的进程，使阻塞的 COM 调用返回
"""
import logging
import os
import threading
import time
from abc import ABC
from contextlib import contextmanager
from dataclasses import dataclass, field
from enum import Enum
from typing import Callable, Dict, Iterable, List, Optional, Set, Tuple

from src.core.models import FileType
from . import process_utils
from .com_utils import (describe_com_error, is_busy_error, is_dead_server_error, is_not_registered_error,
                        safe_get, safe_set)

logger = logging.getLogger("bdp.backend")

LogFunc = Callable[[str], None]


class OfficeKind(Enum):
    """Office 组件类型"""
    WRITER = "writer"              # 文字（Word / WPS 文字）
    SPREADSHEET = "spreadsheet"    # 表格（Excel / WPS 表格）
    PRESENTATION = "presentation"  # 演示（PowerPoint / WPS 演示）


KIND_BY_FILE_TYPE: Dict[FileType, OfficeKind] = {
    FileType.WORD: OfficeKind.WRITER,
    FileType.EXCEL: OfficeKind.SPREADSHEET,
    FileType.PPT: OfficeKind.PRESENTATION,
}


class BackendUnavailableError(RuntimeError):
    """所需的 Office 组件不可用"""


class OfficeStartError(RuntimeError):
    """Office COM 对象创建失败"""


@dataclass
class ComponentStatus:
    """单个组件（如 WPS 表格）的检测结果"""
    backend_key: str
    kind: OfficeKind
    progid: str
    display_name: str
    registered: bool = False
    clsid: str = ""
    server_path: str = ""
    server_exists: bool = False
    error: str = ""
    deep_checked: bool = False
    deep_ok: Optional[bool] = None
    version: str = ""
    deep_detail: str = ""

    @property
    def available(self) -> bool:
        """是否可用：深度检测过则以深度检测为准，否则以注册表为准"""
        if self.deep_checked:
            return bool(self.deep_ok)
        return self.registered and (self.server_exists or not self.server_path)

    def summary(self) -> str:
        mark = "✓" if self.available else "✗"
        text = f"{self.display_name} {mark}"
        if self.version:
            text += f" (v{self.version})"
        return text

    def detail_lines(self) -> List[str]:
        lines = [f"{self.display_name}: {'可用' if self.available else '不可用'}",
                 f"    ProgID: {self.progid}"]
        if self.clsid:
            lines.append(f"    CLSID: {self.clsid}")
        if self.server_path:
            lines.append(f"    LocalServer32: {self.server_path} "
                         f"({'文件存在' if self.server_exists else '文件不存在'})")
        if self.deep_checked:
            lines.append(f"    实际启动测试: {'通过' if self.deep_ok else '失败'} {self.deep_detail}".rstrip())
        if self.error:
            lines.append(f"    错误: {self.error}")
        return lines


def _parse_server_path(command: str) -> str:
    """从 LocalServer32 命令行中取出可执行文件路径"""
    command = (command or "").strip()
    if not command:
        return ""
    if command.startswith('"'):
        end = command.find('"', 1)
        return command[1:end] if end > 0 else command.strip('"')
    lower = command.lower()
    idx = lower.find(".exe")
    if idx >= 0:
        return command[:idx + 4]
    return command.split(" /")[0].split(" -")[0]


def lookup_progid_registration(progid: str) -> Tuple[str, str, str]:
    """
    在注册表中查找 ProgID → CLSID → LocalServer32

    Returns:
        (clsid, server_path, error)
    """
    try:
        import winreg  # type: ignore
    except ImportError:
        return "", "", "非 Windows 系统，无法读取注册表"

    def _read_default(root, subkey, view=0):
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | view) as key:
            value, _ = winreg.QueryValueEx(key, "")
            return value

    clsid = ""
    for candidate in (f"{progid}\\CLSID",):
        try:
            clsid = _read_default(winreg.HKEY_CLASSES_ROOT, candidate)
        except OSError:
            clsid = ""
    if not clsid:
        # 版本无关 ProgID 可能只有 CurVer
        try:
            cur_ver = _read_default(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CurVer")
            clsid = _read_default(winreg.HKEY_CLASSES_ROOT, f"{cur_ver}\\CLSID")
        except OSError:
            return "", "", f"注册表中未找到 {progid}"

    server = ""
    views = [0]
    for name in ("KEY_WOW64_64KEY", "KEY_WOW64_32KEY"):
        if hasattr(winreg, name):
            views.append(getattr(winreg, name))
    for view in views:
        try:
            server = _read_default(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32", view)
            if server:
                break
        except OSError:
            continue
    if not server:
        return clsid, "", f"{progid} 已注册 CLSID {clsid}，但缺少 LocalServer32 注册项"
    return clsid, _parse_server_path(os.path.expandvars(server)), ""


class OfficeBackend(ABC):
    """Office 套件后端基类"""

    key: str = ""
    display_name: str = ""
    PROGIDS: Dict[OfficeKind, str] = {}
    COMPONENT_NAMES: Dict[OfficeKind, str] = {}
    PROCESS_NAMES: Dict[OfficeKind, Tuple[str, ...]] = {}
    # DisplayAlerts 的"全部关闭"取值（各组件不同）
    DISPLAY_ALERTS_OFF: Dict[OfficeKind, object] = {
        OfficeKind.WRITER: 0,          # wdAlertsNone
        OfficeKind.SPREADSHEET: False,
        OfficeKind.PRESENTATION: 1,    # ppAlertsNone
    }

    @property
    def all_process_names(self) -> Set[str]:
        names: Set[str] = set()
        for values in self.PROCESS_NAMES.values():
            names.update(v.lower() for v in values)
        return names

    def progid(self, kind: OfficeKind) -> str:
        return self.PROGIDS[kind]

    def component_name(self, kind: OfficeKind) -> str:
        return self.COMPONENT_NAMES[kind]

    # ---------- 检测 ----------
    def check_component(self, kind: OfficeKind) -> ComponentStatus:
        """快速检测（只读注册表，不启动程序）"""
        progid = self.progid(kind)
        status = ComponentStatus(self.key, kind, progid, self.component_name(kind))
        clsid, server, error = lookup_progid_registration(progid)
        status.clsid = clsid
        status.server_path = server
        status.registered = bool(clsid)
        status.server_exists = bool(server) and os.path.isfile(server)
        if error and not clsid:
            status.error = error
        elif clsid and server and not status.server_exists:
            status.error = f"COM 注册指向的程序不存在: {server}（注册信息损坏，请修复/重装 {self.display_name}）"
        elif error:
            status.error = error
        return status

    def check_all(self) -> Dict[OfficeKind, ComponentStatus]:
        return {kind: self.check_component(kind) for kind in OfficeKind}

    def deep_check_component(self, kind: OfficeKind, timeout: float = 90.0) -> ComponentStatus:
        """
        深度检测：真正创建 COM 对象 → Visible=False → 读取版本 → Quit，并确认进程退出。
        必须在已初始化 COM 的线程中调用。
        """
        status = self.check_component(kind)
        status.deep_checked = True
        session = OfficeSession(self, kind)
        with Watchdog(timeout, lambda: session.kill("深度检测超时"), name=f"deep-{kind.value}") as wd:
            try:
                session.start()
                status.version = str(safe_get(session.app, "Version", "") or "")
                session.close()
                leftovers = [pid for pid in session.spawned_pids if process_utils.pid_exists(pid)]
                status.deep_ok = True
                status.deep_detail = "(创建→隐藏→退出 正常)"
                if leftovers:
                    status.deep_detail += f"，但进程 {leftovers} 未退出"
            except Exception as e:  # noqa: BLE001 - 检测需要捕获所有错误并报告
                status.deep_ok = False
                status.deep_detail = "超时" if wd.fired else ""
                status.error = describe_com_error(e)
                session.kill("深度检测失败清理")
        return status

    # ---------- 会话 ----------
    def open_session(self, kind: OfficeKind, owned_pids: Optional[Set[int]] = None,
                     log: Optional[LogFunc] = None) -> "OfficeSession":
        session = OfficeSession(self, kind, log=log)
        session.start(owned_pids=owned_pids)
        return session

    def open_writer(self, **kwargs) -> "OfficeSession":
        return self.open_session(OfficeKind.WRITER, **kwargs)

    def open_spreadsheet(self, **kwargs) -> "OfficeSession":
        return self.open_session(OfficeKind.SPREADSHEET, **kwargs)

    def open_presentation(self, **kwargs) -> "OfficeSession":
        return self.open_session(OfficeKind.PRESENTATION, **kwargs)

    def __repr__(self):
        return f"<{self.__class__.__name__} {self.display_name}>"


def _create_dispatch(progid: str):
    """创建进程外 COM 对象，并强制使用动态（后期）绑定，避免 gen_py 缓存问题"""
    import pythoncom  # type: ignore
    import pywintypes  # type: ignore
    import win32com.client.dynamic  # type: ignore

    clsid = pywintypes.IID(progid)  # ProgID → CLSID；未注册时抛出 CO_E_CLASSSTRING
    dispatch = pythoncom.CoCreateInstance(
        clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch
    )
    return win32com.client.dynamic.Dispatch(dispatch, progid)


class OfficeSession:
    """一个 Office Application COM 实例的完整生命周期"""

    def __init__(self, backend: OfficeBackend, kind: OfficeKind, log: Optional[LogFunc] = None,
                 dispatch_factory: Optional[Callable[[str], object]] = None):
        self.backend = backend
        self.kind = kind
        self.progid = backend.progid(kind)
        self.app = None
        self.spawned_pids: Set[int] = set()
        self.owned = False           # 进程是否由本程序启动（决定能否 Quit/强制结束）
        self.mode = ""
        self._starting = False
        self._saved_props: Dict[str, object] = {}
        self.jobs_done = 0
        self.dead = False
        self.started_at = 0.0
        self._before: Optional[Dict[int, str]] = None
        self._lock = threading.Lock()
        self._log = log or (lambda msg: logger.info(msg))
        self._dispatch_factory = dispatch_factory or _create_dispatch

    @property
    def label(self) -> str:
        return f"{self.backend.component_name(self.kind)}({self.progid})"

    # 会话与进程的关系：
    #   OWNED        本次创建时出现了新进程 → 完全由本程序管理（隐藏、Quit、超时可强制结束）
    #   POOL_SHARED  没有新进程，但所有候选进程都是本批次启动的（WPS 整合模式下多个组件共用 wps.exe）
    #                → 可以隐藏、超时可结束，但由启动该进程的会话负责 Quit
    #   ATTACHED     复用了用户自己打开的 WPS → 不隐藏窗口、不 Quit、不关闭用户文档、不结束进程，
    #                只临时关闭提示框并在结束时恢复
    OWNED, POOL_SHARED, ATTACHED = "owned", "pool_shared", "attached"

    def start(self, owned_pids: Optional[Set[int]] = None) -> None:
        """创建 COM 对象并进行静默配置"""
        names = self.backend.all_process_names
        with self._lock:
            self._before = process_utils.snapshot(names)
            self._starting = True
        t0 = time.monotonic()
        try:
            app = self._create_with_retry()
        except Exception as e:  # noqa: BLE001
            self.dead = True
            self._kill_new_processes("启动失败清理")
            with self._lock:
                self._starting = False
            if is_not_registered_error(e):
                raise OfficeStartError(
                    f"{self.backend.component_name(self.kind)} COM 接口不可用（ProgID: {self.progid} 未注册）。"
                    f"请安装或修复 {self.backend.display_name}。") from e
            raise OfficeStartError(
                f"无法启动 {self.backend.component_name(self.kind)}（ProgID: {self.progid}）: "
                f"{describe_com_error(e)}") from e
        self.started_at = time.monotonic()

        after = process_utils.snapshot(names)
        new_pids = set(after) - set(self._before or {})
        owned_alive = {pid for pid in (owned_pids or set()) if pid in after}
        with self._lock:
            if new_pids:
                self.mode, self.spawned_pids = self.OWNED, new_pids
            elif not process_utils.tracking_available():
                # 无法识别进程时按"自己启动"处理（与 Word 的 CreateObject 语义一致）
                self.mode = self.OWNED
            elif after and set(after) <= owned_alive:
                self.mode, self.spawned_pids = self.POOL_SHARED, set(after)
            else:
                self.mode = self.ATTACHED
            self.owned = self.mode == self.OWNED
            self.app = app
            self._starting = False
        self._log(f"已启动 {self.label}，耗时 {self.started_at - t0:.1f}s，模式 {self.mode}，"
                  f"进程: {sorted(self.spawned_pids) or '未识别'}")
        if self.mode == self.ATTACHED:
            self._log(f"⚠️ {self.label} 复用了用户已打开的 WPS/Office 实例：不会隐藏窗口、不会退出，"
                      f"超时保护也无法强制结束该进程")
        self._configure()

    def _kill_new_processes(self, reason: str) -> List[int]:
        if self._before is None:
            return []
        pids = process_utils.find_new_pids(self._before, process_names=self.backend.all_process_names)
        return process_utils.kill_pids(pids, reason=f"[{self.label}] {reason}") if pids else []

    def _create_with_retry(self, attempts: int = 3, delay: float = 2.0):
        """
        创建 COM 对象。上一个实例刚被强制结束时，COM 的类工厂登记可能尚未清理，
        CoCreateInstance 会短暂返回 "RPC 服务器不可用"，稍等后重试即可。
        """
        for attempt in range(1, attempts + 1):
            try:
                return self._dispatch_factory(self.progid)
            except Exception as e:  # noqa: BLE001
                if attempt >= attempts or not (is_dead_server_error(e) or is_busy_error(e)):
                    raise
                self._log(f"创建 {self.label} 失败（{describe_com_error(e)}），{delay:.0f} 秒后重试")
                time.sleep(delay)

    def _set(self, name: str, value) -> None:
        """设置属性；复用用户实例时记录原值以便恢复"""
        if self.mode == self.ATTACHED and name not in self._saved_props:
            self._saved_props[name] = safe_get(self.app, name)
        safe_set(self.app, name, value, self._log)

    def _configure(self) -> None:
        if self.mode != self.ATTACHED:
            self._set("Visible", False)
            if self.kind in (OfficeKind.WRITER, OfficeKind.SPREADSHEET):
                self._set("ScreenUpdating", False)
        self._set("DisplayAlerts", self.backend.DISPLAY_ALERTS_OFF[self.kind])
        if self.kind == OfficeKind.SPREADSHEET:
            self._set("AskToUpdateLinks", False)
            self._set("EnableEvents", False)
        # 禁用宏自动运行（msoAutomationSecurityForceDisable = 3）
        self._set("AutomationSecurity", 3)

    def _restore_props(self, app) -> None:
        for name, value in self._saved_props.items():
            if value is not None:
                safe_set(app, name, value, self._log)
        self._saved_props.clear()

    @property
    def is_alive(self) -> bool:
        if self.dead or self.app is None:
            return False
        if self.spawned_pids and not any(process_utils.pid_exists(p) for p in self.spawned_pids):
            return False
        return True

    def kill(self, reason: str = "") -> List[int]:
        """强制结束本会话的进程（可在其他线程调用，只做进程操作不碰 COM）"""
        with self._lock:
            self.dead = True
            pids = set(self.spawned_pids) if self.mode != self.ATTACHED else set()
            starting = self._starting
        if not pids and starting:
            # 仍在创建 COM 对象（例如卡在首次启动的弹窗）：结束新出现的进程
            return self._kill_new_processes(reason)
        if not pids:
            self._log(f"⚠️ {self.label} 无可结束的自有进程（{reason}）")
            return []
        return process_utils.kill_pids(pids, reason=f"[{self.label}] {reason}")

    def close(self, quit_timeout: float = 15.0) -> None:
        """退出应用并确保进程结束；可重复调用。Quit 本身卡住时由看门狗强制结束进程"""
        with self._lock:
            app, self.app = self.app, None
        if app is not None and not self.dead:
            with Watchdog(quit_timeout + 15, lambda: self.kill("关闭/退出超时"), name=f"close-{self.kind.value}"):
                try:
                    if self.mode == self.OWNED:
                        # 关闭可能残留的（本程序打开的）文档，不保存
                        self._close_leftover_documents(app)
                        app.Quit()
                    elif self.mode == self.ATTACHED:
                        self._restore_props(app)
                except Exception as e:  # noqa: BLE001 - 退出失败时走强制结束
                    self._log(f"{self.label} 退出失败: {describe_com_error(e)}")
        del app
        try:
            import pythoncom  # type: ignore
            pythoncom.CoFreeUnusedLibraries()
        except Exception:  # noqa: BLE001 - 非 Windows / 测试环境
            pass
        if self.mode == self.OWNED and self.spawned_pids:
            leftovers = process_utils.wait_for_exit(self.spawned_pids, quit_timeout)
            if leftovers:
                process_utils.kill_pids(leftovers, reason=f"[{self.label}] Quit 后 {quit_timeout:.0f}s 仍未退出")
        self.dead = True

    def _close_leftover_documents(self, app) -> None:
        collection_name = {OfficeKind.WRITER: "Documents", OfficeKind.SPREADSHEET: "Workbooks",
                           OfficeKind.PRESENTATION: "Presentations"}[self.kind]
        collection = safe_get(app, collection_name)
        count = safe_get(collection, "Count", 0) if collection is not None else 0
        for _ in range(int(count or 0)):
            try:
                doc = collection.Item(1)
                if self.kind == OfficeKind.WRITER:
                    doc.Close(0)
                elif self.kind == OfficeKind.SPREADSHEET:
                    doc.Close(False)
                else:
                    doc.Close()
            except Exception:  # noqa: BLE001 - 尽力而为
                break


class OfficeSessionPool:
    """
    批量打印期间复用 Office 实例，减少反复启动 WPS 的开销。
    必须在同一个（已 CoInitialize 的）线程中调用 get/discard/close_all；
    kill_kind 可在看门狗线程中调用。
    """

    def __init__(self, selection: "BackendSelection", log: Optional[LogFunc] = None,
                 recycle_after: int = 30, session_factory=None):
        self.selection = selection
        self.recycle_after = recycle_after
        self._sessions: Dict[OfficeKind, OfficeSession] = {}
        self._owned_pids: Set[int] = set()
        self._lock = threading.Lock()
        self._log = log or (lambda msg: logger.info(msg))
        self._session_factory = session_factory or (lambda backend, kind: OfficeSession(backend, kind, log=self._log))
        self.all_spawned_pids: Set[int] = set()

    def backend_for(self, kind: OfficeKind) -> OfficeBackend:
        backend, reason = self.selection.backend_for(kind)
        if backend is None:
            raise BackendUnavailableError(reason)
        return backend

    def get(self, kind: OfficeKind) -> OfficeSession:
        with self._lock:
            session = self._sessions.get(kind)
        if session is not None:
            if session.is_alive and session.jobs_done < self.recycle_after:
                return session
            reason = "达到复用上限，重新启动" if session.is_alive else "实例已失效，重新启动"
            self._log(f"{session.label} {reason}")
            self.discard(kind)
        backend = self.backend_for(kind)
        session = self._session_factory(backend, kind)
        with self._lock:
            self._sessions[kind] = session
        with self._lock:
            # 去掉已退出的进程，避免 PID 被系统复用后误认
            self._owned_pids = {pid for pid in self._owned_pids if process_utils.pid_exists(pid)}
            owned = set(self._owned_pids)
        session.start(owned_pids=owned)
        with self._lock:
            if session.owned:
                self._owned_pids |= session.spawned_pids
                self.all_spawned_pids |= session.spawned_pids
        return session

    def peek(self, kind: OfficeKind) -> Optional[OfficeSession]:
        with self._lock:
            return self._sessions.get(kind)

    def discard(self, kind: OfficeKind) -> None:
        with self._lock:
            session = self._sessions.pop(kind, None)
        if session is not None:
            try:
                session.close()
            except Exception as e:  # noqa: BLE001
                self._log(f"关闭 {session.label} 失败: {describe_com_error(e)}")
                session.kill("关闭失败")

    def kill_kind(self, kind: OfficeKind, reason: str) -> List[int]:
        """看门狗回调：结束该组件会话的进程，并把共享同一进程的其他会话标记为失效"""
        with self._lock:
            session = self._sessions.get(kind)
            others = [s for k, s in self._sessions.items() if k != kind]
        if session is None:
            return []
        killed = session.kill(reason)
        for other in others:
            if set(killed) & other.spawned_pids:
                other.dead = True
        return killed

    def close_all(self) -> None:
        for kind in list(OfficeKind):
            self.discard(kind)
        # 最终兜底：本批次启动过的进程一个都不留
        leftovers = [pid for pid in self.all_spawned_pids if process_utils.pid_exists(pid)]
        if leftovers:
            process_utils.kill_pids(leftovers, reason="批次结束兜底清理")


class OfficeBusyError(RuntimeError):
    """另一个任务（打印 / 页数统计 / 深度检测）正在使用 WPS/Office 自动化"""


# 进程内同一时间只允许一个任务驱动 WPS/Office：多个会话池同时运行时，
# "进程快照差集"无法区分各自启动的进程，可能误结束对方的 WPS
OFFICE_AUTOMATION_LOCK = threading.Lock()


@contextmanager
def office_automation(timeout: float, busy_message: str):
    if not OFFICE_AUTOMATION_LOCK.acquire(timeout=max(0.0, timeout)):
        raise OfficeBusyError(busy_message)
    try:
        yield
    finally:
        OFFICE_AUTOMATION_LOCK.release()


class Watchdog:
    """超时看门狗：在 with 块执行超过 timeout 秒时调用 on_timeout（在独立线程中）"""

    def __init__(self, timeout: float, on_timeout: Callable[[], object], name: str = "watchdog"):
        self.timeout = timeout
        self.on_timeout = on_timeout
        self.name = name
        self.fired = False
        self._timer: Optional[threading.Timer] = None

    def _fire(self):
        self.fired = True
        logger.error("看门狗 %s 触发：操作超过 %.0f 秒", self.name, self.timeout)
        try:
            self.on_timeout()
        except Exception:  # noqa: BLE001 - 看门狗线程不能抛异常
            logger.exception("看门狗回调失败")

    def __enter__(self):
        if self.timeout and self.timeout > 0:
            self._timer = threading.Timer(self.timeout, self._fire)
            self._timer.daemon = True
            self._timer.name = f"Watchdog-{self.name}"
            self._timer.start()
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._timer is not None:
            self._timer.cancel()
        return False


@contextmanager
def com_initialized():
    """在当前线程初始化 COM（STA），退出时反初始化"""
    try:
        import pythoncom  # type: ignore
    except ImportError:
        yield
        return
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass


@dataclass
class BackendSelection:
    """每个组件实际使用的后端"""
    preference: str = "auto"
    chosen: Dict[OfficeKind, Optional[OfficeBackend]] = field(default_factory=dict)
    statuses: Dict[str, Dict[OfficeKind, ComponentStatus]] = field(default_factory=dict)
    messages: Dict[OfficeKind, str] = field(default_factory=dict)

    def backend_for(self, kind: OfficeKind) -> Tuple[Optional[OfficeBackend], str]:
        return self.chosen.get(kind), self.messages.get(kind, "")

    @property
    def primary_display(self) -> str:
        names = {b.display_name for b in self.chosen.values() if b is not None}
        if not names:
            return "未检测到 WPS Office / Microsoft Office"
        return " + ".join(sorted(names))

    @property
    def any_available(self) -> bool:
        return any(b is not None for b in self.chosen.values())


def select_backends(backends: Iterable[OfficeBackend], preference: str = "auto",
                    statuses: Optional[Dict[str, Dict[OfficeKind, ComponentStatus]]] = None) -> BackendSelection:
    """
    按偏好为每个组件选择后端。

    auto: 每个组件优先 WPS，WPS 该组件不可用时回退 MS Office
    wps / msoffice: 只使用指定套件
    """
    backends = list(backends)
    selection = BackendSelection(preference=preference)
    statuses = statuses or {b.key: b.check_all() for b in backends}
    selection.statuses = statuses

    if preference in ("wps", "msoffice"):
        ordered = [b for b in backends if b.key == preference]
    else:
        ordered = sorted(backends, key=lambda b: 0 if b.key == "wps" else 1)

    for kind in OfficeKind:
        chosen = None
        problems = []
        for backend in ordered:
            status = statuses.get(backend.key, {}).get(kind)
            if status is not None and status.available:
                chosen = backend
                break
            if status is not None:
                problems.append(f"{status.display_name}（ProgID: {status.progid}）: {status.error or '不可用'}")
        selection.chosen[kind] = chosen
        if chosen is None:
            wps_status = statuses.get("wps", {}).get(kind)
            wps_installed = any(s.registered for s in statuses.get("wps", {}).values())
            if wps_installed and wps_status is not None and not wps_status.available:
                head = f"检测到 WPS Office，但{wps_status.display_name} COM 接口不可用。ProgID: {wps_status.progid}"
            else:
                head = "未找到可用的 Office 组件，请安装 WPS Office（推荐）或 Microsoft Office"
            selection.messages[kind] = head + ("\n" + "\n".join(problems) if problems else "")
        else:
            selection.messages[kind] = f"使用 {chosen.display_name}"
    return selection


def default_backends() -> List[OfficeBackend]:
    from .wps_backend import WpsBackend
    from .ms_office_backend import MsOfficeBackend
    return [WpsBackend(), MsOfficeBackend()]


def detect_backend(preference: str = "auto") -> BackendSelection:
    """检测环境并返回组件 → 后端的选择结果（只读注册表，不启动程序）"""
    return select_backends(default_backends(), preference)
