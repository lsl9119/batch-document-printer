"""
Microsoft Office 后端（可选回退）

仅在 WPS 对应组件不可用、且用户未限定"仅 WPS"时使用。
"""
from .office_backend import OfficeBackend, OfficeKind


class MsOfficeBackend(OfficeBackend):
    key = "msoffice"
    display_name = "Microsoft Office"
    PROGIDS = {
        OfficeKind.WRITER: "Word.Application",
        OfficeKind.SPREADSHEET: "Excel.Application",
        OfficeKind.PRESENTATION: "PowerPoint.Application",
    }
    COMPONENT_NAMES = {
        OfficeKind.WRITER: "Microsoft Word",
        OfficeKind.SPREADSHEET: "Microsoft Excel",
        OfficeKind.PRESENTATION: "Microsoft PowerPoint",
    }
    PROCESS_NAMES = {
        OfficeKind.WRITER: ("winword.exe",),
        OfficeKind.SPREADSHEET: ("excel.exe",),
        OfficeKind.PRESENTATION: ("powerpnt.exe",),
    }
