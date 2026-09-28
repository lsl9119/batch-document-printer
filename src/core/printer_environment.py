"""
打印机环境管理（替代 v5.x 的 PrinterConfigManager）

批量打印期间临时修改两类系统状态，并保证恢复：

1. 每用户默认打印参数（PRINTER_INFO_9 / 每用户 DEVMODE）：纸张、方向、双面、彩色
   - v5.x 使用 SetPrinter level 2（全局设置），普通用户没有权限，实际会静默失败；
     level 9 只影响当前用户，无需管理员权限。
2. Windows 默认打印机：WPS 演示的 PrintOptions.ActivePrinter 不一定可写，WPS 文字/表格
   设置 ActivePrinter 的行为也因版本而异。批次包含 Office 文档且目标打印机不是默认
   打印机时，临时把目标打印机设为默认，结束后恢复。

安全保证：
- 修改前先写入恢复状态文件（data/printer_restore_state.json）
- with 语句 finally 中恢复；atexit 兜底
- 程序崩溃/被强制结束后，下次启动时 recover_pending_state() 自动恢复
- 全局锁保证同一时间只有一个批次修改打印机环境
"""
import atexit
import json
import logging
import os
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional

from .models import ColorMode, DuplexMode, Orientation, PrintSettings

logger = logging.getLogger("bdp.printer_env")

LogFunc = Callable[[str], None]

# DEVMODE 常量
DM_ORIENTATION = 0x00000001
DM_PAPERSIZE = 0x00000002
DM_COPIES = 0x00000100
DM_COLOR = 0x00000800
DM_DUPLEX = 0x00001000
DM_COLLATE = 0x00008000
DM_IN_BUFFER = 8
DM_OUT_BUFFER = 2
DMORIENT_PORTRAIT, DMORIENT_LANDSCAPE = 1, 2
DMCOLOR_MONOCHROME, DMCOLOR_COLOR = 1, 2
DMDUP_SIMPLEX, DMDUP_VERTICAL, DMDUP_HORIZONTAL = 1, 2, 3  # VERTICAL=长边翻转, HORIZONTAL=短边翻转

# DeviceCapabilities 常量
DC_PAPERS = 2
DC_DUPLEX = 7
DC_PAPERNAMES = 16
DC_ORIENTATION = 17
DC_COLORDEVICE = 32

PRINTER_ACCESS_USE = 0x00000008

# 标准纸张 → DMPAPER_*
STANDARD_PAPER_CODES: Dict[str, int] = {
    "Letter": 1, "Tabloid": 3, "Legal": 5, "A3": 8, "A4": 9, "A5": 11,
    "B4": 12, "B5": 13, "A6": 70,
}

UNSUPPORTED_MSG = "当前打印机驱动不支持该设置，已使用驱动默认值"

_ENV_LOCK = threading.Lock()  # 非可重入锁：允许在 atexit（主线程）中释放
_active_environment: Optional["PrinterEnvironment"] = None


def _state_file_default() -> Path:
    from src.utils.path_utils import get_config_dir
    return get_config_dir() / "printer_restore_state.json"


def _win32print():
    import win32print  # type: ignore
    return win32print


# ---------------------------------------------------------------------------
# 默认打印机
# ---------------------------------------------------------------------------

def get_default_printer() -> Optional[str]:
    try:
        return _win32print().GetDefaultPrinter() or None
    except Exception as e:  # noqa: BLE001 - 没有默认打印机时会抛异常
        logger.warning("获取默认打印机失败: %s", e)
        return None


def set_default_printer(name: str) -> bool:
    wp = _win32print()
    try:
        wp.SetDefaultPrinter(name)
    except Exception as e:  # noqa: BLE001
        logger.error("设置默认打印机为 %s 失败: %s", name, e)
        return False
    # 部分系统上 SetDefaultPrinter 生效有延迟
    for _ in range(10):
        if (get_default_printer() or "").lower() == name.lower():
            return True
        time.sleep(0.2)
    return (get_default_printer() or "").lower() == name.lower()


# ---------------------------------------------------------------------------
# 能力检测
# ---------------------------------------------------------------------------

@dataclass
class PrinterCapabilities:
    printer_name: str
    supports_duplex: Optional[bool] = None
    supports_color: Optional[bool] = None
    supports_landscape: Optional[bool] = None
    paper_codes: List[int] = field(default_factory=list)
    paper_names: List[str] = field(default_factory=list)

    def paper_code_for(self, paper_name: str) -> Optional[int]:
        """纸张名 → DMPAPER 代码：先查驱动返回的名称，再查标准表"""
        if not paper_name:
            return None
        lowered = paper_name.strip().lower()
        for code, name in zip(self.paper_codes, self.paper_names):
            if name.strip().lower() == lowered:
                return code
        return STANDARD_PAPER_CODES.get(paper_name) or STANDARD_PAPER_CODES.get(paper_name.upper())


def get_printer_port(printer_name: str) -> str:
    wp = _win32print()
    handle = wp.OpenPrinter(printer_name, {"DesiredAccess": PRINTER_ACCESS_USE})
    try:
        return wp.GetPrinter(handle, 2).get("pPortName", "") or ""
    finally:
        wp.ClosePrinter(handle)


def get_capabilities(printer_name: str) -> PrinterCapabilities:
    caps = PrinterCapabilities(printer_name)
    try:
        wp = _win32print()
        port = get_printer_port(printer_name)
    except Exception as e:  # noqa: BLE001
        logger.warning("无法打开打印机 %s 读取能力: %s", printer_name, e)
        return caps

    def _cap(index):
        try:
            return wp.DeviceCapabilities(printer_name, port, index)
        except Exception as e:  # noqa: BLE001 - 部分驱动不实现某些能力查询
            logger.debug("DeviceCapabilities(%s, %s) 失败: %s", printer_name, index, e)
            return None

    duplex = _cap(DC_DUPLEX)
    caps.supports_duplex = None if duplex is None or duplex < 0 else bool(duplex)
    color = _cap(DC_COLORDEVICE)
    caps.supports_color = None if color is None or color < 0 else bool(color)
    orientation = _cap(DC_ORIENTATION)
    caps.supports_landscape = None if orientation is None or orientation < 0 else orientation != 0
    papers = _cap(DC_PAPERS)
    names = _cap(DC_PAPERNAMES)
    if papers:
        caps.paper_codes = [int(p) for p in papers]
    if names:
        caps.paper_names = [str(n).strip("\x00 ").strip() for n in names]
    return caps


# ---------------------------------------------------------------------------
# 每用户 DEVMODE
# ---------------------------------------------------------------------------

DEVMODE_FIELDS = ("Duplex", "Color", "Orientation", "PaperSize", "Copies", "Collate", "Fields")


def _devmode_snapshot(devmode) -> Dict[str, int]:
    return {name: int(getattr(devmode, name)) for name in DEVMODE_FIELDS if hasattr(devmode, name)}


def build_devmode_plan(settings: PrintSettings, caps: PrinterCapabilities):
    """
    根据打印设置和驱动能力计算要写入的 DEVMODE 字段。

    Returns:
        (fields: {字段: 值}, field_mask, warnings)
    """
    values: Dict[str, int] = {}
    mask = 0
    warnings: List[str] = []

    duplex = settings.effective_duplex
    if duplex == DuplexMode.SIMPLEX:
        if caps.supports_duplex is not False:
            values["Duplex"] = DMDUP_SIMPLEX
            mask |= DM_DUPLEX
    elif caps.supports_duplex is False:
        warnings.append(f"双面打印：{UNSUPPORTED_MSG}")
    else:
        values["Duplex"] = DMDUP_VERTICAL if duplex == DuplexMode.DUPLEX_LONG else DMDUP_HORIZONTAL
        mask |= DM_DUPLEX

    if settings.color_mode == ColorMode.COLOR and caps.supports_color is False:
        warnings.append(f"彩色打印：{UNSUPPORTED_MSG}（该打印机为黑白设备）")
    else:
        values["Color"] = DMCOLOR_COLOR if settings.color_mode == ColorMode.COLOR else DMCOLOR_MONOCHROME
        mask |= DM_COLOR

    if settings.orientation != Orientation.AUTO:
        if settings.orientation == Orientation.LANDSCAPE and caps.supports_landscape is False:
            warnings.append(f"横向打印：{UNSUPPORTED_MSG}")
        else:
            values["Orientation"] = (DMORIENT_LANDSCAPE if settings.orientation == Orientation.LANDSCAPE
                                     else DMORIENT_PORTRAIT)
            mask |= DM_ORIENTATION

    code = caps.paper_code_for(settings.paper_size)
    if code is None:
        warnings.append(f"纸张 {settings.paper_size}：无法识别该纸张，已使用驱动默认值")
    elif caps.paper_codes and code not in caps.paper_codes:
        warnings.append(f"纸张 {settings.paper_size}：{UNSUPPORTED_MSG}")
    else:
        values["PaperSize"] = code
        mask |= DM_PAPERSIZE

    # 份数由应用程序控制，驱动级固定为 1，避免份数被放大
    values["Copies"] = 1
    values["Collate"] = 1
    mask |= DM_COPIES | DM_COLLATE
    return values, mask, warnings


class DevModeGuard:
    """修改指定打印机的每用户默认 DEVMODE，并可恢复"""

    def __init__(self, printer_name: str, log: LogFunc):
        self.printer_name = printer_name
        self._log = log
        self.had_user_devmode: Optional[bool] = None
        self.original: Dict[str, int] = {}
        self.applied = False

    def _open(self):
        wp = _win32print()
        return wp, wp.OpenPrinter(self.printer_name, {"DesiredAccess": PRINTER_ACCESS_USE})

    def backup(self) -> None:
        wp, handle = self._open()
        try:
            user_devmode = None
            try:
                user_devmode = wp.GetPrinter(handle, 9).get("pDevMode")
            except Exception as e:  # noqa: BLE001 - 旧驱动/旧 pywin32 不支持 level 9
                self._log(f"读取每用户打印参数失败: {e}")
            self.had_user_devmode = user_devmode is not None
            base = user_devmode or wp.GetPrinter(handle, 2).get("pDevMode")
            self.original = _devmode_snapshot(base) if base is not None else {}
        finally:
            wp.ClosePrinter(handle)

    def apply(self, settings: PrintSettings, caps: PrinterCapabilities) -> List[str]:
        values, mask, warnings = build_devmode_plan(settings, caps)
        wp, handle = self._open()
        try:
            devmode = None
            try:
                devmode = wp.GetPrinter(handle, 9).get("pDevMode")
            except Exception:  # noqa: BLE001
                devmode = None
            if devmode is None:
                devmode = wp.GetPrinter(handle, 2).get("pDevMode")
            if devmode is None:
                warnings.append("无法读取打印机默认参数（DEVMODE），纸张/方向/双面/色彩将使用驱动默认值")
                return warnings
            for name, value in values.items():
                setattr(devmode, name, value)
            devmode.Fields = int(devmode.Fields) | mask
            try:
                wp.DocumentProperties(0, handle, self.printer_name, devmode, devmode,
                                      DM_IN_BUFFER | DM_OUT_BUFFER)
            except Exception as e:  # noqa: BLE001 - 部分驱动不支持校验，继续写入
                self._log(f"驱动校验打印参数失败（继续）: {e}")
            wp.SetPrinter(handle, 9, {"pDevMode": devmode}, 0)
            self.applied = True
            # 读回验证：驱动可能静默拒绝某些字段
            try:
                check = wp.GetPrinter(handle, 9).get("pDevMode")
            except Exception:  # noqa: BLE001
                check = None
            if check is not None:
                labels = {"Duplex": "双面", "Color": "色彩", "Orientation": "方向", "PaperSize": "纸张"}
                for name, label in labels.items():
                    if name in values and int(getattr(check, name, values[name])) != values[name]:
                        warnings.append(f"{label}：驱动未接受该设置（{UNSUPPORTED_MSG}）")
        except Exception as e:  # noqa: BLE001 - 设置失败不应中断打印
            warnings.append(f"写入每用户打印参数失败（{e}），纸张/方向/双面/色彩将使用驱动默认值")
        finally:
            wp.ClosePrinter(handle)
        return warnings

    def restore(self) -> bool:
        if not self.applied:
            return True
        return restore_devmode(self.printer_name, bool(self.had_user_devmode), self.original, self._log)


def clear_user_devmode_native(printer_name: str) -> bool:
    """
    删除每用户默认 DEVMODE（SetPrinter level 9, pDevMode=NULL），恢复为跟随打印机全局设置。
    pywin32 的 SetPrinter 不接受 pDevMode=None，因此直接调用 winspool.drv。
    """
    import ctypes
    from ctypes import wintypes

    class PRINTER_DEFAULTSW(ctypes.Structure):
        _fields_ = [("pDatatype", wintypes.LPWSTR), ("pDevMode", ctypes.c_void_p),
                    ("DesiredAccess", wintypes.DWORD)]

    class PRINTER_INFO_9W(ctypes.Structure):
        _fields_ = [("pDevMode", ctypes.c_void_p)]

    winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)
    winspool.OpenPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.HANDLE),
                                      ctypes.POINTER(PRINTER_DEFAULTSW)]
    winspool.SetPrinterW.argtypes = [wintypes.HANDLE, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD]
    winspool.ClosePrinter.argtypes = [wintypes.HANDLE]
    handle = wintypes.HANDLE()
    defaults = PRINTER_DEFAULTSW(None, None, PRINTER_ACCESS_USE)
    if not winspool.OpenPrinterW(printer_name, ctypes.byref(handle), ctypes.byref(defaults)):
        return False
    try:
        info = PRINTER_INFO_9W(None)
        return bool(winspool.SetPrinterW(handle, 9, ctypes.byref(info), 0))
    finally:
        winspool.ClosePrinter(handle)


def restore_devmode(printer_name: str, had_user_devmode: bool, original: Dict[str, int], log: LogFunc) -> bool:
    if not had_user_devmode:
        try:
            if clear_user_devmode_native(printer_name):
                log(f"已清除临时写入的每用户打印参数: {printer_name}")
                return True
        except (OSError, AttributeError, ValueError) as e:  # 非 Windows / 旧系统：回退为逐字段恢复
            log(f"清除每用户打印参数失败，改为逐字段恢复: {e}")
    wp = _win32print()
    try:
        handle = wp.OpenPrinter(printer_name, {"DesiredAccess": PRINTER_ACCESS_USE})
    except Exception as e:  # noqa: BLE001
        log(f"恢复打印参数失败，无法打开打印机 {printer_name}: {e}")
        return False
    try:
        devmode = None
        try:
            devmode = wp.GetPrinter(handle, 9).get("pDevMode")
        except Exception:  # noqa: BLE001
            pass
        if devmode is None:
            return True
        for name, value in original.items():
            try:
                setattr(devmode, name, value)
            except Exception:  # noqa: BLE001 - 个别字段不可写时跳过
                pass
        wp.SetPrinter(handle, 9, {"pDevMode": devmode}, 0)
        log(f"已恢复打印机原始参数: {printer_name}")
        return True
    except Exception as e:  # noqa: BLE001
        log(f"恢复打印机参数失败 {printer_name}: {e}")
        return False
    finally:
        wp.ClosePrinter(handle)


# ---------------------------------------------------------------------------
# 批次环境
# ---------------------------------------------------------------------------

class PrinterEnvironmentBusy(RuntimeError):
    """另一个批次/实例正在修改打印机环境"""


# ---------------------------------------------------------------------------
# "让 Windows 管理默认打印机"（LegacyDefaultPrinterMode）
# ---------------------------------------------------------------------------
_WINDOWS_KEY = r"Software\Microsoft\Windows NT\CurrentVersion\Windows"
_LEGACY_VALUE = "LegacyDefaultPrinterMode"
_MISSING = "__missing__"


def get_legacy_default_mode():
    """返回 LegacyDefaultPrinterMode 的值；不存在返回 _MISSING；无法读取返回 None"""
    try:
        import winreg  # type: ignore
    except ImportError:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _WINDOWS_KEY) as key:
            value, _ = winreg.QueryValueEx(key, _LEGACY_VALUE)
            return int(value)
    except FileNotFoundError:
        return _MISSING
    except OSError:
        return None


def set_legacy_default_mode(value) -> bool:
    """恢复 LegacyDefaultPrinterMode（_MISSING 表示删除该值）"""
    if value is None:
        return True
    try:
        import winreg  # type: ignore
    except ImportError:
        return True
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _WINDOWS_KEY, 0,
                            winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE) as key:
            if value == _MISSING:
                try:
                    winreg.DeleteValue(key, _LEGACY_VALUE)
                except FileNotFoundError:
                    pass
            else:
                winreg.SetValueEx(key, _LEGACY_VALUE, 0, winreg.REG_DWORD, int(value))
        return True
    except OSError as e:
        logger.warning("恢复“让 Windows 管理默认打印机”设置失败: %s", e)
        return False


# ---------------------------------------------------------------------------
# 进程身份（防止 PID 被复用后误判“另一个实例正在打印”）
# ---------------------------------------------------------------------------
def _process_create_time(pid: int) -> Optional[float]:
    try:
        import psutil  # type: ignore
        return float(psutil.Process(pid).create_time())
    except Exception:  # noqa: BLE001 - 进程不存在或无法读取
        return None


def _pid_alive(pid: int, create_time: Optional[float] = None) -> bool:
    try:
        import psutil  # type: ignore
    except ImportError:
        return False
    try:
        if not psutil.pid_exists(pid):
            return False
    except Exception:  # noqa: BLE001
        return False
    if create_time is None:
        return True
    actual = _process_create_time(pid)
    return actual is None or abs(actual - create_time) < 2.0


def _state_owner_alive(state: dict) -> bool:
    pid = int(state.get("pid") or 0)
    if not pid or pid == os.getpid():
        return False
    return _pid_alive(pid, state.get("pid_create_time"))


def _load_state(state_file: Path) -> Optional[dict]:
    try:
        return json.loads(state_file.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return None
    except (OSError, ValueError):
        return {}


def _state_devmodes(state: dict) -> List[dict]:
    """兼容旧格式（单个 devmode）与新格式（devmodes 列表）"""
    entries = list(state.get("devmodes") or [])
    if state.get("devmode"):
        entries.append(state["devmode"])
    return [e for e in entries if e and e.get("printer")]


class PrinterEnvironment:
    """
    with PrinterEnvironment(printer, settings, switch_default=True) as env:
        ... 打印 ...
    # 离开 with 时恢复默认打印机和打印参数
    """

    def __init__(self, printer_name: str, settings: PrintSettings, switch_default: bool,
                 apply_devmode: bool = True, state_file: Optional[Path] = None,
                 log: Optional[LogFunc] = None):
        self.printer_name = printer_name
        self.settings = settings
        self.switch_default = switch_default
        self.apply_devmode = apply_devmode
        self.state_file = state_file or _state_file_default()
        self._log = log or (lambda m: logger.info(m))
        self.original_default: Optional[str] = None
        self.legacy_mode = None
        self.default_switched = False
        self.devmode_guard: Optional[DevModeGuard] = None
        self.warnings: List[str] = []
        self.capabilities: Optional[PrinterCapabilities] = None
        # 上次异常退出且未能恢复的状态：继承下来，保证最终恢复到“真正的原始状态”
        self._inherited: Optional[dict] = None
        self._inherited_devmodes: List[dict] = []
        self._owns_state_file = False
        self._restored = False
        self._lock_acquired = False

    # ---- 状态文件 ----
    def _devmode_entries(self) -> List[dict]:
        entries = list(self._inherited_devmodes)
        guard = self.devmode_guard
        if guard is not None and guard.had_user_devmode is not None:
            entries.append({"printer": self.printer_name, "had_user_devmode": guard.had_user_devmode,
                            "original": guard.original})
        return entries

    def _write_state(self) -> None:
        state = {
            "pid": os.getpid(),
            "pid_create_time": _process_create_time(os.getpid()),
            "time": time.strftime("%Y-%m-%d %H:%M:%S"),
            "target_printer": self.printer_name,
            "original_default": self.original_default,
            "default_switched": self.default_switched,
            "legacy_default_mode": self.legacy_mode,
            "devmodes": self._devmode_entries(),
        }
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.state_file.with_suffix(".tmp")
        tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self.state_file)
        self._owns_state_file = True

    def _check_foreign_state(self) -> None:
        state = _load_state(self.state_file)
        if state is None:
            return
        if _state_owner_alive(state):
            raise PrinterEnvironmentBusy("另一个批量打印程序实例正在打印，请等待其完成后再试")
        # 残留的旧状态（上次异常退出）：先尝试恢复
        recover_pending_state(self.state_file, self._log)
        remaining = _load_state(self.state_file)
        if remaining:
            # 恢复失败：继承旧状态，绝不能用当前（仍被修改的）状态覆盖“原始状态”
            self._log("⚠️ 上次遗留的打印机状态未能完全恢复，本次结束时将一并恢复")
            self._inherited = remaining

    # ---- 进入/退出 ----
    def __enter__(self):
        global _active_environment
        if not _ENV_LOCK.acquire(timeout=5):
            raise PrinterEnvironmentBusy("已有批量打印任务正在修改打印机设置")
        self._lock_acquired = True
        try:
            self._check_foreign_state()
            inherited = self._inherited or {}
            if inherited.get("default_switched") and inherited.get("original_default"):
                self.original_default = inherited["original_default"]
                self.default_switched = True  # 结束时需要恢复到继承的原默认打印机
                self.legacy_mode = inherited.get("legacy_default_mode")
            else:
                self.original_default = get_default_printer()
            stale_devmodes = _state_devmodes(inherited)
            same_printer = [e for e in stale_devmodes if e["printer"] == self.printer_name]
            self._inherited_devmodes = [e for e in stale_devmodes if e["printer"] != self.printer_name]

            if self.apply_devmode:
                self.capabilities = get_capabilities(self.printer_name)
                self.devmode_guard = DevModeGuard(self.printer_name, self._log)
                if same_printer:
                    # 该打印机的参数上次未恢复：以上次备份的原始值为准
                    self.devmode_guard.had_user_devmode = bool(same_printer[0].get("had_user_devmode"))
                    self.devmode_guard.original = same_printer[0].get("original") or {}
                    self.devmode_guard.applied = True
                else:
                    try:
                        self.devmode_guard.backup()
                    except Exception as e:  # noqa: BLE001
                        self._log(f"备份打印机参数失败，跳过参数设置: {e}")
                        self.devmode_guard = None
            # 先落盘“原始状态”，再做任何修改
            self._write_state()
            _active_environment = self

            if self.devmode_guard is not None:
                self.warnings.extend(self.devmode_guard.apply(self.settings, self.capabilities))

            if self.switch_default and self.printer_name and \
                    (get_default_printer() or "").lower() != self.printer_name.lower():
                if not self.default_switched:
                    self.legacy_mode = get_legacy_default_mode()
                self.default_switched = True
                self._write_state()
                if set_default_printer(self.printer_name):
                    self._log(f"已临时将系统默认打印机切换为: {self.printer_name}"
                              f"（原默认: {self.original_default or '无'}）")
                else:
                    self.warnings.append(f"无法将默认打印机临时切换为 {self.printer_name}，"
                                         f"Office 文档可能打印到原默认打印机 {self.original_default}")
            for warning in self.warnings:
                self._log(f"⚠️ {warning}")
            return self
        except BaseException:
            self.restore()
            raise

    def __exit__(self, exc_type, exc, tb):
        self.restore()
        return False

    def restore(self) -> bool:
        global _active_environment
        if self._restored:
            return True
        ok = True
        try:
            if self.default_switched and self.original_default:
                current = get_default_printer() or ""
                if current.lower() != self.original_default.lower():
                    if set_default_printer(self.original_default):
                        self._log(f"已恢复系统默认打印机: {self.original_default}")
                    elif not _printer_exists(self.original_default):
                        self._log(f"原默认打印机 {self.original_default} 已不存在，无法恢复")
                    else:
                        ok = False
                        self._log(f"❌ 恢复默认打印机失败: {self.original_default}（下次启动时将重试）")
                set_legacy_default_mode(self.legacy_mode)
            if self.devmode_guard is not None:
                ok = self.devmode_guard.restore() and ok
            for entry in self._inherited_devmodes:
                ok = restore_devmode(entry["printer"], bool(entry.get("had_user_devmode")),
                                     entry.get("original") or {}, self._log) and ok
            if ok and (self._owns_state_file or self._inherited):
                try:
                    self.state_file.unlink()
                except FileNotFoundError:
                    pass
                except OSError as e:
                    self._log(f"删除恢复状态文件失败: {e}")
            self._restored = True
            return ok
        finally:
            if _active_environment is self:
                _active_environment = None
            if self._lock_acquired:
                self._lock_acquired = False
                _ENV_LOCK.release()


def _printer_exists(name: str) -> bool:
    try:
        wp = _win32print()
        flags = wp.PRINTER_ENUM_LOCAL | wp.PRINTER_ENUM_CONNECTIONS
        return any(p[2].lower() == name.lower() for p in wp.EnumPrinters(flags))
    except Exception:  # noqa: BLE001 - 无法枚举时按存在处理（保守：保留状态文件）
        return True


def recover_pending_state(state_file: Optional[Path] = None, log: Optional[LogFunc] = None) -> List[str]:
    """
    启动时调用：若上次运行异常退出导致默认打印机/打印参数未恢复，则在此恢复。
    恢复失败时保留状态文件，下次继续尝试。

    Returns:
        执行的恢复操作说明列表
    """
    state_file = state_file or _state_file_default()
    log = log or (lambda m: logger.info(m))
    actions: List[str] = []
    state = _load_state(state_file)
    if state is None:
        return actions
    if state == {}:
        log("恢复状态文件损坏，已忽略")
        try:
            state_file.unlink()
        except OSError:
            pass
        return actions
    if _state_owner_alive(state):
        log(f"另一个实例 (pid={state.get('pid')}) 正在打印，暂不恢复打印机状态")
        return actions

    ok = True
    original = state.get("original_default")
    if state.get("default_switched") and original:
        current = get_default_printer()
        if (current or "").lower() != original.lower():
            if set_default_printer(original):
                actions.append(f"已恢复上次异常退出前的默认打印机: {original}")
            elif not _printer_exists(original):
                actions.append(f"上次的原默认打印机 {original} 已不存在，无法恢复")
            else:
                ok = False
                actions.append(f"恢复默认打印机失败: {original}")
        set_legacy_default_mode(state.get("legacy_default_mode"))
    for entry in _state_devmodes(state):
        if restore_devmode(entry["printer"], bool(entry.get("had_user_devmode")),
                           entry.get("original") or {}, log):
            actions.append(f"已恢复打印机参数: {entry['printer']}")
        elif _printer_exists(entry["printer"]):
            ok = False
    if ok:
        try:
            state_file.unlink()
        except OSError:
            pass
    for action in actions:
        log(action)
    return actions


def _atexit_restore():
    env = _active_environment
    if env is not None:
        logger.warning("程序退出时打印机环境仍处于修改状态，执行恢复")
        try:
            env.restore()
        except Exception:  # noqa: BLE001
            logger.exception("退出时恢复打印机环境失败")


atexit.register(_atexit_restore)
