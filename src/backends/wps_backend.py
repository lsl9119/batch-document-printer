"""
WPS Office 后端

COM ProgID（参考 harness-anything 项目 cli_anything/wps/utils/wps_backend.py 的调用方式，MIT 许可）:
  KWPS.Application  → WPS 文字
  KET.Application   → WPS 表格
  KWPP.Application  → WPS 演示

WPS 的对象模型与 Microsoft Office 高度兼容，差异（命名参数支持、
PrintOptions.ActivePrinter 可写性等）在 office_ops 中通过降级调用处理。
"""
from .office_backend import OfficeBackend, OfficeKind


class WpsBackend(OfficeBackend):
    key = "wps"
    display_name = "WPS Office"
    PROGIDS = {
        OfficeKind.WRITER: "KWPS.Application",
        OfficeKind.SPREADSHEET: "KET.Application",
        OfficeKind.PRESENTATION: "KWPP.Application",
    }
    COMPONENT_NAMES = {
        OfficeKind.WRITER: "WPS 文字",
        OfficeKind.SPREADSHEET: "WPS 表格",
        OfficeKind.PRESENTATION: "WPS 演示",
    }
    # WPS 2019+ 整合模式下各组件可能都运行在 wps.exe 中，因此统一跟踪三个进程名
    PROCESS_NAMES = {
        OfficeKind.WRITER: ("wps.exe", "et.exe", "wpp.exe"),
        OfficeKind.SPREADSHEET: ("wps.exe", "et.exe", "wpp.exe"),
        OfficeKind.PRESENTATION: ("wps.exe", "et.exe", "wpp.exe"),
    }
