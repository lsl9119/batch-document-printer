"""
文档处理器基础类
定义所有文档处理器必须实现的接口
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional, Set

from ..core.models import FileType, PrintSettings
from ..core.print_context import PrintContext, PrintJobError  # noqa: F401 - 供子类导入


class BaseDocumentHandler(ABC):
    """文档处理器基础抽象类"""

    #: 是否需要 WPS / Microsoft Office
    requires_office: bool = False

    @abstractmethod
    def get_supported_file_types(self) -> Set[FileType]:
        """获取支持的文件类型"""

    @abstractmethod
    def get_supported_extensions(self) -> Set[str]:
        """获取支持的文件扩展名（小写，包含点号，如：{'.pdf', '.docx'}）"""

    def can_handle_file(self, file_path: Path) -> bool:
        """检查是否能处理指定文件（默认：文件存在且扩展名匹配）"""
        if not self.validate_file_exists(file_path):
            return False
        return file_path.suffix.lower() in self.get_supported_extensions()

    @abstractmethod
    def print_document(self, file_path: Path, settings: PrintSettings,
                       context: Optional[PrintContext] = None) -> bool:
        """
        打印文档

        Returns:
            True 表示打印作业已成功提交

        Raises:
            PrintJobError: 打印失败，异常信息为面向用户的原因说明
        """

    @abstractmethod
    def count_pages(self, file_path: Path, context: Optional[PrintContext] = None) -> int:
        """统计文档页数（失败时抛出异常）"""

    def get_handler_name(self) -> str:
        return self.__class__.__name__

    def describe_backend(self, context: Optional[PrintContext] = None) -> str:
        """日志中显示的打印引擎名称（优先使用处理器在本次任务中记录的实际引擎）"""
        if context is not None and context.backend_used:
            return context.backend_used
        return self.get_handler_name()

    def validate_file_exists(self, file_path: Path) -> bool:
        return file_path.exists() and file_path.is_file()

    def get_file_size_mb(self, file_path: Path) -> float:
        try:
            return file_path.stat().st_size / (1024 * 1024)
        except OSError:
            return 0.0
