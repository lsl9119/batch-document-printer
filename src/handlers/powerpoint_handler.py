"""
演示文稿处理器（PPT / PPTX / DPS）

通过 WPS 演示（KWPP.Application）打印。WPS 演示的 PrintOptions.ActivePrinter
在部分版本中不可写，因此批量打印时由 PrintController 临时把目标打印机设为
Windows 默认打印机（finally 中恢复），这里仍会尝试设置并在日志中记录是否可写。
"""
from pathlib import Path
from typing import Optional

from ..backends.office_backend import OfficeKind, OfficeSession
from ..backends import office_ops as ops
from ..core.models import FileType, PrintSettings
from ..core.print_context import PrintContext
from .office_handler import OfficeDocumentHandler


class PowerPointDocumentHandler(OfficeDocumentHandler):
    """演示文稿处理器"""

    KIND = OfficeKind.PRESENTATION
    FILE_TYPE = FileType.PPT
    EXTENSIONS = {'.ppt', '.pptx', '.dps'}

    def get_handler_name(self) -> str:
        return "演示文稿处理器"

    def _print_with_session(self, session: OfficeSession, path: Path, settings: PrintSettings,
                            ctx: PrintContext) -> None:
        pres = None
        try:
            pres = ops.open_presentation(session.app, path, ctx.log)
            result = ops.presentation_print(pres, settings, ctx.log)
            ctx.notes.append(result["method"])
            if not result["active_printer_writable"]:
                if ctx.default_printer_switched:
                    ctx.notes.append("PrintOptions.ActivePrinter 不可写，已通过临时切换默认打印机指定目标打印机")
                else:
                    ctx.notes.append("PrintOptions.ActivePrinter 不可写，使用系统默认打印机（即目标打印机）")
        finally:
            ops.close_presentation(pres, ctx.log)

    def _count_with_session(self, session: OfficeSession, path: Path, ctx: PrintContext) -> int:
        pres = None
        try:
            pres = ops.open_presentation(session.app, path, ctx.log)
            return ops.presentation_count_slides(pres, ctx.log)
        finally:
            ops.close_presentation(pres, ctx.log)

    def count_pages(self, file_path: Path, context: Optional[PrintContext] = None) -> int:
        """.pptx 优先用 python-pptx 离线统计（无需启动 WPS），失败再走 COM"""
        if file_path.suffix.lower() == '.pptx':
            try:
                from pptx import Presentation  # type: ignore
                return len(Presentation(str(file_path)).slides)
            except Exception:  # noqa: BLE001 - 回退 COM
                pass
        return super().count_pages(file_path, context)
