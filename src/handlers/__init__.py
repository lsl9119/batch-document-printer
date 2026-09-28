"""
文档处理器模块
提供基于文件格式的模块化文档处理能力
"""
from .base_handler import BaseDocumentHandler
from .handler_registry import HandlerRegistry, create_default_registry
from .pdf_handler import PDFDocumentHandler
from .word_handler import WordDocumentHandler
from .powerpoint_handler import PowerPointDocumentHandler
from .excel_handler import ExcelDocumentHandler
from .image_handler import ImageDocumentHandler
from .text_handler import TextDocumentHandler

__all__ = [
    'BaseDocumentHandler',
    'HandlerRegistry',
    'create_default_registry',
    'PDFDocumentHandler',
    'WordDocumentHandler',
    'PowerPointDocumentHandler',
    'ExcelDocumentHandler',
    'ImageDocumentHandler',
    'TextDocumentHandler'
]
