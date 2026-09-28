"""
文字文档处理器（DOC / DOCX / WPS）
通过 Office 后端调用 WPS 文字（KWPS.Application），无 WPS 时可回退 Microsoft Word
"""
from pathlib import Path

from ..backends.office_backend import OfficeKind, OfficeSession
from ..backends import office_ops as ops
from ..core.models import FileType, PrintSettings
from ..core.print_context import PrintContext
from .office_handler import OfficeDocumentHandler


class WordDocumentHandler(OfficeDocumentHandler):
    """文字文档处理器"""

    KIND = OfficeKind.WRITER
    FILE_TYPE = FileType.WORD
    EXTENSIONS = {'.doc', '.docx', '.wps'}

    def get_handler_name(self) -> str:
        return "文字文档处理器"

    def _print_with_session(self, session: OfficeSession, path: Path, settings: PrintSettings,
                            ctx: PrintContext) -> None:
        app = session.app
        ops.try_set_active_printer(app, settings.printer_name, ctx.log, "Application")
        doc = None
        try:
            doc = ops.open_writer_document(app, path, ctx.log)
            if settings.office_force_page_setup:
                ctx.notes.extend(ops.apply_writer_page_setup(doc, settings, ctx.log))
            method = ops.writer_print(app, doc, settings.copies, ctx.log)
            ctx.notes.append(method)
        finally:
            ops.close_writer_document(doc, ctx.log)

    def _count_with_session(self, session: OfficeSession, path: Path, ctx: PrintContext) -> int:
        doc = None
        try:
            doc = ops.open_writer_document(session.app, path, ctx.log)
            return ops.writer_count_pages(doc, ctx.log)
        finally:
            ops.close_writer_document(doc, ctx.log)
