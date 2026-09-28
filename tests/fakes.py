"""
测试替身：win32print 与 WPS/Office COM 对象

单元测试与 CI 不接触真实打印机和真实 WPS，全部通过这些替身验证逻辑。
"""
import threading
from typing import Dict, List, Optional

from src.backends import com_utils


# ---------------------------------------------------------------------------
# COM 异常
# ---------------------------------------------------------------------------
class FakeComError(Exception):
    """模拟 pywintypes.com_error: args = (hresult, strerror, excepinfo, argerror)"""

    def __init__(self, hresult: int, text: str = "", description: str = "", scode: int = 0):
        excepinfo = (0, "FakeApp", description, None, 0, scode) if (description or scode) else None
        super().__init__(hresult, text, excepinfo, None)
        self.hresult = hresult


def signature_error():
    return FakeComError(com_utils.DISP_E_UNKNOWNNAME, "未知名称")


def dead_server_error():
    return FakeComError(com_utils.RPC_S_SERVER_UNAVAILABLE, "RPC 服务器不可用")


def password_error():
    return FakeComError(com_utils.DISP_E_EXCEPTION, "发生意外", "密码不正确，无法打开文档", -2146822880)


# ---------------------------------------------------------------------------
# COM 对象
# ---------------------------------------------------------------------------
class _NamedArgsGuard:
    named_args_supported = True

    def _check(self, kwargs):
        if kwargs and not self.named_args_supported:
            raise signature_error()


class FakePageSetup:
    def __init__(self, orientation=1, paper=9, print_area=""):
        self.Orientation = orientation
        self.PaperSize = paper
        self.PrintArea = print_area


class FakeWriterDoc(_NamedArgsGuard):
    def __init__(self, app, path, pages=3):
        self.app = app
        self.path = path
        self.pages = pages
        self.closed = False
        self.printouts: List[dict] = []
        self.PageSetup = FakePageSetup(orientation=0, paper=7)
        self.named_args_supported = app.named_args_supported

    def PrintOut(self, *args, **kwargs):
        self._check(kwargs)
        if self.app.hang_event is not None:
            self.app.hang_event.wait(30)
            if self.app.killed:
                raise dead_server_error()
        if self.app.print_error is not None:
            raise self.app.print_error
        self.printouts.append({"args": args, "kwargs": kwargs, "printer": self.app.ActivePrinter})
        self.app.printed.append((self.path, args, kwargs))

    def Close(self, *args, **kwargs):
        self._check(kwargs)
        self.closed = True
        self.app.open_docs.remove(self)

    def ComputeStatistics(self, what):
        return self.pages


class FakeCollection(_NamedArgsGuard):
    def __init__(self, app, factory):
        self.app = app
        self.factory = factory
        self.named_args_supported = app.named_args_supported

    def Open(self, *args, **kwargs):
        self._check(kwargs)
        path = kwargs.get("FileName") or kwargs.get("Filename") or args[0]
        if self.app.open_error is not None:
            raise self.app.open_error
        self.app.open_calls.append((args, kwargs))
        doc = self.factory(self.app, path)
        self.app.open_docs.append(doc)
        return doc

    @property
    def Count(self):
        return len(self.app.open_docs)

    def Item(self, index):
        return self.app.open_docs[index - 1]


class FakeOfficeApp:
    """通用 Application 替身；kind 决定集合名称"""

    def __init__(self, progid: str, named_args_supported: bool = True, active_printer: str = "Default Printer",
                 active_printer_writable: bool = True):
        self.progid = progid
        self.named_args_supported = named_args_supported
        self.Visible = True
        self.DisplayAlerts = True
        self._active_printer = active_printer
        self.active_printer_writable = active_printer_writable
        self.Version = "12.1.0.fake"
        self.BackgroundPrintingStatus = 0
        self.quit_called = False
        self.killed = False
        self.printed: List[tuple] = []
        self.open_calls: List[tuple] = []
        self.open_docs: List[object] = []
        self.open_error: Optional[Exception] = None
        self.print_error: Optional[Exception] = None
        self.hang_event: Optional[threading.Event] = None
        self.Documents = FakeCollection(self, lambda app, p: FakeWriterDoc(app, p))
        self.Workbooks = FakeCollection(self, lambda app, p: FakeWorkbook(app, p))
        self.Presentations = FakeCollection(self, lambda app, p: FakePresentation(app, p))

    @property
    def ActivePrinter(self):
        return self._active_printer

    @ActivePrinter.setter
    def ActivePrinter(self, value):
        if not self.active_printer_writable:
            raise FakeComError(com_utils.DISP_E_EXCEPTION, "只读属性", "ActivePrinter 不可写")
        self._active_printer = value

    def Quit(self):
        self.quit_called = True


class FakeWorksheet(_NamedArgsGuard):
    def __init__(self, wb, name, visible=-1, pages=1, has_content=True, print_area=""):
        self.wb = wb
        self.Name = name
        self.Visible = visible
        self._pages = pages
        self.PageSetup = FakePageSetup(print_area=print_area)
        self.PageSetup.Pages = type("Pages", (), {"Count": pages})()
        self._has_content = has_content
        self.UsedRange = type("R", (), {"Address": "$A$1:$D$40" if has_content else "$A$1"})()
        self.Shapes = type("S", (), {"Count": 0})()
        self.printed = 0

    def Range(self, address):
        return type("Cell", (), {"Value": "x" if self._has_content else None})()

    def PrintOut(self, *args, **kwargs):
        self.printed += 1
        self.wb.app.printed.append((self.wb.path, "sheet", self.Name))


class FakeSheets:
    def __init__(self, sheets):
        self._sheets = sheets

    @property
    def Count(self):
        return len(self._sheets)

    def Item(self, i):
        return self._sheets[i - 1]


class FakeWorkbook(_NamedArgsGuard):
    def __init__(self, app, path, sheets=None):
        self.app = app
        self.path = path
        self.named_args_supported = app.named_args_supported
        specs = getattr(app, "sheet_specs", None) or [dict(name="Sheet1", pages=2)]
        self.Worksheets = FakeSheets([FakeWorksheet(self, **spec) for spec in specs])
        self.Charts = type("C", (), {"Count": 0})()
        self.printouts: List[dict] = []

    def PrintOut(self, *args, **kwargs):
        self._check(kwargs)
        if self.app.print_error is not None:
            raise self.app.print_error
        self.printouts.append({"args": args, "kwargs": kwargs})
        self.app.printed.append((self.path, args, kwargs))

    def Close(self, *args, **kwargs):
        self._check(kwargs)
        self.app.open_docs.remove(self)


class FakePrintOptions:
    def __init__(self, app):
        self.app = app
        self.PrintInBackground = -1
        self.NumberOfCopies = 1
        self.Collate = -1
        self.PrintColorType = 1
        self.FitToPage = 0
        self.OutputType = 1
        self.RangeType = 1
        self._active_printer = app.ActivePrinter
        self.copies_writable = True

    @property
    def ActivePrinter(self):
        return self._active_printer

    @ActivePrinter.setter
    def ActivePrinter(self, value):
        if not self.app.active_printer_writable:
            raise FakeComError(com_utils.DISP_E_EXCEPTION, "只读属性", "ActivePrinter 属性为只读")
        self._active_printer = value

    def __setattr__(self, key, value):
        if key == "NumberOfCopies" and getattr(self, "copies_writable", True) is False:
            raise signature_error()
        object.__setattr__(self, key, value)


class FakePresentation(_NamedArgsGuard):
    def __init__(self, app, path):
        self.app = app
        self.path = path
        self.named_args_supported = app.named_args_supported
        self.PrintOptions = FakePrintOptions(app)
        self.Slides = type("Slides", (), {"Count": 5})()
        self.Saved = 0

    def PrintOut(self, *args, **kwargs):
        self._check(kwargs)
        if self.app.print_error is not None:
            raise self.app.print_error
        self.app.printed.append((self.path, args, kwargs, self.PrintOptions.NumberOfCopies))

    def Close(self):
        self.app.open_docs.remove(self)


# ---------------------------------------------------------------------------
# win32print
# ---------------------------------------------------------------------------
class FakeDevMode:
    def __init__(self, **fields):
        self.Duplex = 1
        self.Color = 2
        self.Orientation = 1
        self.PaperSize = 9
        self.Copies = 1
        self.Collate = 0
        self.Fields = 0
        for k, v in fields.items():
            setattr(self, k, v)

    def copy(self):
        return FakeDevMode(**{k: getattr(self, k) for k in ("Duplex", "Color", "Orientation",
                                                                "PaperSize", "Copies", "Collate", "Fields")})


class FakePrinter:
    def __init__(self, name, port="PORTPROMPT:", duplex=True, color=True, landscape=True,
                 papers=(9, 8, 11, 1), paper_names=("A4", "A3", "A5", "Letter")):
        self.name = name
        self.port = port
        self.global_devmode = FakeDevMode()
        self.user_devmode: Optional[FakeDevMode] = None
        self.caps = {7: 1 if duplex else 0, 32: 1 if color else 0, 17: 90 if landscape else 0,
                     2: list(papers), 16: list(paper_names)}


class FakeWin32Print:
    PRINTER_ENUM_LOCAL = 2
    PRINTER_ENUM_CONNECTIONS = 4
    PRINTER_ACCESS_USE = 8

    def __init__(self, printers: List[FakePrinter], default: Optional[str]):
        self.printers: Dict[str, FakePrinter] = {p.name: p for p in printers}
        self.default = default
        self.open_handles = 0
        self.set_default_calls: List[str] = []
        self.fail_set_default = False

    # 默认打印机
    def GetDefaultPrinter(self):
        if not self.default:
            raise RuntimeError("no default printer")
        return self.default

    def SetDefaultPrinter(self, name):
        if self.fail_set_default:
            raise RuntimeError("SetDefaultPrinter failed")
        if name not in self.printers:
            raise RuntimeError(f"unknown printer {name}")
        self.set_default_calls.append(name)
        self.default = name

    def EnumPrinters(self, flags, *args):
        return [(0, "", name, "") for name in self.printers]

    # 句柄
    def OpenPrinter(self, name, defaults=None):
        if name not in self.printers:
            raise RuntimeError(f"OpenPrinter failed: {name}")
        self.open_handles += 1
        return self.printers[name]

    def ClosePrinter(self, handle):
        self.open_handles -= 1

    def GetPrinter(self, handle, level):
        if level == 2:
            return {"pPrinterName": handle.name, "pPortName": handle.port, "pDriverName": "FakeDriver",
                    "pDevMode": handle.global_devmode.copy(), "Status": 0}
        if level == 9:
            return {"pDevMode": handle.user_devmode.copy() if handle.user_devmode else None}
        raise RuntimeError("unsupported level")

    def SetPrinter(self, handle, level, info, command):
        assert level == 9, "只允许写每用户 DEVMODE（level 9），不需要管理员权限"
        devmode = info.get("pDevMode")
        handle.user_devmode = devmode.copy() if devmode is not None else None

    def DocumentProperties(self, hwnd, handle, name, out_dm, in_dm, mode):
        return 1

    def DeviceCapabilities(self, name, port, capability, devmode=None):
        return self.printers[name].caps.get(capability)
