"""
图片文档处理器（JPG/JPEG/PNG/BMP/TIFF/WEBP）

主方案：SumatraPDF（支持全部打印参数）；备用：Windows GDI 直接打印（Pillow）。
不再使用"用默认程序打开"这类实际上并未打印却返回成功的方案。
"""
from pathlib import Path
from typing import Any, Dict, Optional, Set

from ..core.models import FileType, PrintSettings
from ..core.print_context import PrintContext, PrintJobError
from .base_handler import BaseDocumentHandler
from . import gdi_print, sumatra

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:  # pragma: no cover
    PIL_AVAILABLE = False


class ImageDocumentHandler(BaseDocumentHandler):
    """图片文档处理器"""

    _EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}

    def get_handler_name(self) -> str:
        return "图片处理器"

    def describe_backend(self, context: Optional[PrintContext] = None) -> str:
        if context is not None and context.backend_used:
            return context.backend_used
        return "SumatraPDF" if sumatra.sumatra_available() else "Windows GDI"

    def get_supported_file_types(self) -> Set[FileType]:
        return {FileType.IMAGE}

    def get_supported_extensions(self) -> Set[str]:
        return set(self._EXTENSIONS)

    def can_handle_file(self, file_path: Path) -> bool:
        if not super().can_handle_file(file_path):
            return False
        try:
            if file_path.stat().st_size < 10:
                return False
        except OSError:
            return False
        if PIL_AVAILABLE:
            try:
                with Image.open(file_path) as img:
                    img.verify()
            except Exception:  # noqa: BLE001 - 损坏的图片
                return False
        return True

    def count_pages(self, file_path: Path, context: Optional[PrintContext] = None) -> int:
        if not self.can_handle_file(file_path):
            raise RuntimeError("文件已损坏或不是有效的图片")
        if file_path.suffix.lower() in ('.tiff', '.tif') and PIL_AVAILABLE:
            with Image.open(file_path) as img:
                return max(1, int(getattr(img, "n_frames", 1)))
        return 1

    def print_document(self, file_path: Path, settings: PrintSettings,
                       context: Optional[PrintContext] = None) -> bool:
        context = context or PrintContext()
        if not self.can_handle_file(file_path):
            raise PrintJobError("文件已损坏或不是有效的图片")
        printer = settings.printer_name
        if not printer:
            raise PrintJobError("未设置打印机")
        sumatra_error = None
        if sumatra.sumatra_available():
            try:
                used = sumatra.run_print(file_path, printer, settings, context.log)
                context.backend_used = "SumatraPDF"
                context.notes.append(f"print-settings={used}")
                return True
            except PrintJobError as e:
                sumatra_error = e
                context.log(f"SumatraPDF 打印图片失败，改用 GDI 备用方案: {e}")
        try:
            gdi_print.print_image(printer, file_path, settings, context.log)
        except ImportError as e:
            raise PrintJobError(f"GDI 打印组件不可用: {e}" +
                                (f"；SumatraPDF 错误: {sumatra_error}" if sumatra_error else "")) from e
        context.backend_used = "Windows GDI"
        return True

    def get_file_info(self, file_path: Path) -> Dict[str, Any]:
        info: Dict[str, Any] = {
            'file_path': str(file_path), 'file_name': file_path.name,
            'file_size': file_path.stat().st_size, 'format': file_path.suffix.upper().lstrip('.'),
            'dimensions': None, 'color_mode': None, 'handler': self.get_handler_name(),
        }
        if PIL_AVAILABLE:
            try:
                with Image.open(file_path) as img:
                    info['dimensions'] = f"{img.width} x {img.height}"
                    info['color_mode'] = img.mode
                    info['format'] = img.format or info['format']
            except Exception:  # noqa: BLE001
                pass
        return info
