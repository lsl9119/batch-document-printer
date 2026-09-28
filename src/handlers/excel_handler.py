"""
表格文档处理器（XLS / XLSX / ET）

直接通过 COM 控制 WPS 表格（KET.Application），不再依赖 xlwings
（xlwings 面向 Microsoft Excel，在只安装 WPS 的环境中不可用）。
"""
from pathlib import Path

from ..backends.office_backend import OfficeKind, OfficeSession
from ..backends import office_ops as ops
from ..core.models import FileType, PrintSettings
from ..core.print_context import PrintContext
from .office_handler import OfficeDocumentHandler


class ExcelDocumentHandler(OfficeDocumentHandler):
    """表格文档处理器"""

    KIND = OfficeKind.SPREADSHEET
    FILE_TYPE = FileType.EXCEL
    EXTENSIONS = {'.xls', '.xlsx', '.et'}

    def get_handler_name(self) -> str:
        return "表格文档处理器"

    def _print_with_session(self, session: OfficeSession, path: Path, settings: PrintSettings,
                            ctx: PrintContext) -> None:
        app = session.app
        ops.try_set_active_printer(app, settings.printer_name, ctx.log, "Application")
        wb = None
        try:
            wb = ops.open_workbook(app, path, ctx.log)
            if settings.office_force_page_setup:
                ctx.notes.extend(ops.apply_spreadsheet_page_setup(wb, settings, ctx.log))
            method = ops.spreadsheet_print(app, wb, settings.copies, ctx.log)
            ctx.notes.append(method)
        finally:
            ops.close_workbook(wb, ctx.log)

    def _count_with_session(self, session: OfficeSession, path: Path, ctx: PrintContext) -> int:
        wb = None
        try:
            wb = ops.open_workbook(session.app, path, ctx.log)
            return ops.spreadsheet_count_pages(wb, ctx.log)
        finally:
            ops.close_workbook(wb, ctx.log)
