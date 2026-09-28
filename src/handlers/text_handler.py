"""
文本文档处理器（TXT）

主方案：Windows GDI 直接打印（可指定打印机、份数，中文编码自动识别）
备用方案：notepad /pt 打印到指定打印机
"""
import os
import subprocess
from pathlib import Path
from typing import Any, Dict, Optional, Set, Tuple

from ..core.models import FileType, PrintSettings
from ..core.print_context import PrintContext, PrintJobError
from .base_handler import BaseDocumentHandler
from . import gdi_print, sumatra

ENCODINGS = ('utf-8-sig', 'utf-16', 'gb18030', 'big5', 'latin1')
MAX_TEXT_MB = 100


def _is_likely_text(content: str) -> bool:
    if not content:
        return True
    printable = sum(1 for c in content if c.isprintable() or c.isspace())
    return printable / len(content) >= 0.8


def read_text_file(file_path: Path) -> Tuple[str, str]:
    """读取文本文件，返回 (内容, 编码)"""
    data = file_path.read_bytes()
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        return data.decode("utf-16"), "utf-16"
    for encoding in ENCODINGS:
        if encoding == 'utf-16':
            continue
        try:
            text = data.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue
        if _is_likely_text(text[:2000]):
            return text, encoding
    raise ValueError(f"无法识别文件编码或疑似二进制文件: {file_path.name}")


class TextDocumentHandler(BaseDocumentHandler):
    """文本文档处理器"""

    _chars_per_line = 75
    _lines_per_page = 50

    def get_handler_name(self) -> str:
        return "文本处理器"

    def describe_backend(self, context: Optional[PrintContext] = None) -> str:
        if context is not None and context.backend_used:
            return context.backend_used
        return "Windows GDI"

    def get_supported_file_types(self) -> Set[FileType]:
        return {FileType.TEXT}

    def get_supported_extensions(self) -> Set[str]:
        return {'.txt'}

    def can_handle_file(self, file_path: Path) -> bool:
        if not super().can_handle_file(file_path):
            return False
        if self.get_file_size_mb(file_path) > MAX_TEXT_MB:
            return False
        try:
            read_text_file(file_path)
            return True
        except (OSError, ValueError):
            return False

    def count_pages(self, file_path: Path, context: Optional[PrintContext] = None) -> int:
        """估算页数（与 GDI 打印版式接近：每行约 75 字符，每页 50 行）"""
        text, _ = read_text_file(file_path)
        pages = gdi_print.layout_text(text, len, self._chars_per_line, self._lines_per_page)
        return max(1, len(pages))

    def print_document(self, file_path: Path, settings: PrintSettings,
                       context: Optional[PrintContext] = None) -> bool:
        context = context or PrintContext()
        printer = settings.printer_name
        if not printer:
            raise PrintJobError("未设置打印机")
        try:
            text, encoding = read_text_file(file_path)
        except (OSError, ValueError) as e:
            raise PrintJobError(str(e)) from e
        try:
            pages = gdi_print.print_text(printer, file_path.name, text, settings.copies, context.log)
            context.backend_used = "Windows GDI"
            context.notes.append(f"编码 {encoding}，每份 {pages} 页")
            return True
        except ImportError as e:
            context.log(f"GDI 组件不可用（{e}），改用记事本打印")
        except PrintJobError as e:
            context.log(f"GDI 文本打印失败（{e}），改用记事本打印")
        return self._print_with_notepad(file_path, printer, settings.copies, context)

    def _print_with_notepad(self, file_path: Path, printer: str, copies: int, context: PrintContext) -> bool:
        if os.name != 'nt':
            raise PrintJobError("文本打印仅支持 Windows")
        for _ in range(max(1, copies)):
            try:
                returncode, _detail = sumatra.run_detached(['notepad.exe', '/pt', str(file_path), printer], 120)
            except subprocess.TimeoutExpired as e:
                raise PrintJobError("记事本打印超时") from e
            except OSError as e:
                raise PrintJobError(f"无法启动记事本: {e}") from e
            if returncode != 0:
                raise PrintJobError(f"记事本打印失败（退出码 {returncode}）")
        context.backend_used = "Notepad"
        return True

    def get_file_info(self, file_path: Path) -> Dict[str, Any]:
        info: Dict[str, Any] = {'file_path': str(file_path), 'file_name': file_path.name,
                                'file_size': file_path.stat().st_size, 'format': 'TXT',
                                'encoding': 'Unknown', 'lines': 0, 'pages': 0}
        try:
            text, encoding = read_text_file(file_path)
            info.update(encoding=encoding.upper(), lines=len(text.split('\n')),
                        pages=self.count_pages(file_path))
        except (OSError, ValueError):
            pass
        return info
