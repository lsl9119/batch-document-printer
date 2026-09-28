"""
文档处理器注册中心
负责管理和分发各种文档格式的处理器
"""
import logging
from pathlib import Path
from typing import Dict, List, Optional, Set

from ..core.models import FileType
from .base_handler import BaseDocumentHandler

logger = logging.getLogger("bdp.handlers")


class HandlerRegistry:
    """处理器注册中心"""

    def __init__(self):
        self._handlers: List[BaseDocumentHandler] = []
        self._file_type_handlers: Dict[FileType, BaseDocumentHandler] = {}
        self._extension_handlers: Dict[str, BaseDocumentHandler] = {}

    def register_handler(self, handler: BaseDocumentHandler):
        if handler in self._handlers:
            return
        self._handlers.append(handler)
        for file_type in handler.get_supported_file_types():
            if file_type in self._file_type_handlers:
                logger.warning("文件类型 %s 的处理器 %s 被 %s 替换", file_type.value,
                               self._file_type_handlers[file_type].get_handler_name(),
                               handler.get_handler_name())
            self._file_type_handlers[file_type] = handler
        for extension in handler.get_supported_extensions():
            self._extension_handlers[extension.lower()] = handler
        logger.debug("已注册处理器 %s: %s", handler.get_handler_name(),
                     sorted(handler.get_supported_extensions()))

    def unregister_handler(self, handler: BaseDocumentHandler):
        if handler not in self._handlers:
            return
        self._handlers.remove(handler)
        for file_type in handler.get_supported_file_types():
            if self._file_type_handlers.get(file_type) is handler:
                del self._file_type_handlers[file_type]
        for extension in handler.get_supported_extensions():
            if self._extension_handlers.get(extension.lower()) is handler:
                del self._extension_handlers[extension.lower()]

    def get_handler_by_file_type(self, file_type: FileType) -> Optional[BaseDocumentHandler]:
        return self._file_type_handlers.get(file_type)

    def get_handler_by_extension(self, extension: str) -> Optional[BaseDocumentHandler]:
        if not extension.startswith('.'):
            extension = '.' + extension
        return self._extension_handlers.get(extension.lower())

    def get_handler_by_file_path(self, file_path: Path) -> Optional[BaseDocumentHandler]:
        return self.get_handler_by_extension(file_path.suffix.lower())

    def can_handle_file(self, file_path: Path) -> bool:
        handler = self.get_handler_by_file_path(file_path)
        return bool(handler and handler.can_handle_file(file_path))

    def get_all_supported_extensions(self) -> Set[str]:
        return set(self._extension_handlers.keys())

    def get_all_supported_file_types(self) -> Set[FileType]:
        return set(self._file_type_handlers.keys())

    def get_registered_handlers(self) -> List[BaseDocumentHandler]:
        return self._handlers.copy()

    def print_registry_info(self):
        for handler in self._handlers:
            logger.info("处理器 %s: %s", handler.get_handler_name(), sorted(handler.get_supported_extensions()))


def create_default_registry() -> HandlerRegistry:
    """创建包含全部内置处理器的注册中心"""
    from .pdf_handler import PDFDocumentHandler
    from .word_handler import WordDocumentHandler
    from .powerpoint_handler import PowerPointDocumentHandler
    from .excel_handler import ExcelDocumentHandler
    from .image_handler import ImageDocumentHandler
    from .text_handler import TextDocumentHandler

    registry = HandlerRegistry()
    for handler in (PDFDocumentHandler(), WordDocumentHandler(), PowerPointDocumentHandler(),
                    ExcelDocumentHandler(), ImageDocumentHandler(), TextDocumentHandler()):
        registry.register_handler(handler)
    return registry
