"""
PDF 文档处理器 —— 基于随程序分发的 SumatraPDF（无需 WPS/Office）

支持：目标打印机、份数、纸张、单/双面（长边/短边）、方向、彩色/黑白、缩放
"""
from pathlib import Path
from typing import Optional, Set

from ..core.models import FileType, PrintSettings
from ..core.print_context import PrintContext, PrintJobError
from .base_handler import BaseDocumentHandler
from . import sumatra


class PDFDocumentHandler(BaseDocumentHandler):
    """PDF 文档处理器"""

    def get_handler_name(self) -> str:
        return "PDF处理器"

    def describe_backend(self, context: Optional[PrintContext] = None) -> str:
        return "SumatraPDF"

    def ensure_sumatra_available(self) -> bool:
        return sumatra.sumatra_available()

    def get_supported_file_types(self) -> Set[FileType]:
        return {FileType.PDF}

    def get_supported_extensions(self) -> Set[str]:
        return {'.pdf'}

    def print_document(self, file_path: Path, settings: PrintSettings,
                       context: Optional[PrintContext] = None) -> bool:
        context = context or PrintContext()
        printer = settings.printer_name or _default_printer()
        if not printer:
            raise PrintJobError("未设置打印机，且系统没有默认打印机")
        used = sumatra.run_print(file_path, printer, settings, context.log)
        context.notes.append(f"print-settings={used}")
        return True

    def count_pages(self, file_path: Path, context: Optional[PrintContext] = None) -> int:
        try:
            from PyPDF2 import PdfReader  # type: ignore
        except ImportError as e:
            raise RuntimeError("需要安装 PyPDF2 来统计 PDF 页数") from e
        try:
            with open(file_path, 'rb') as f:
                reader = PdfReader(f)
                if reader.is_encrypted:
                    try:
                        reader.decrypt("")
                    except Exception as e:  # noqa: BLE001
                        raise RuntimeError("文件被加密") from e
                return len(reader.pages)
        except RuntimeError:
            raise
        except Exception as e:  # noqa: BLE001
            raise RuntimeError(f"文件已损坏: {e}") from e


def _default_printer() -> str:
    from ..core.printer_environment import get_default_printer
    return get_default_printer() or ""
