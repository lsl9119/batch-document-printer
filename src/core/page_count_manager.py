"""
页数统计管理器（WPS 版）

Office 文档通过 Office 后端（WPS 优先）统计，所有文档在同一个后台线程中顺序处理，
复用 WPS 实例并为每个文档设置超时看门狗。
"""
import logging
import time
from typing import List, Optional, Callable, Tuple
from dataclasses import dataclass, field
from enum import Enum

from .models import Document, FileType, OFFICE_FILE_TYPES
from .print_context import PrintContext
from contextlib import nullcontext

from ..backends.office_backend import (KIND_BY_FILE_TYPE, OfficeSessionPool, Watchdog,
                                       com_initialized, detect_backend, office_automation)
from ..handlers import create_default_registry

logger = logging.getLogger("bdp.pagecount")

PAGE_COUNT_TIMEOUT = 120


class PageCountStatus(Enum):
    """页数统计状态"""
    UNKNOWN = "unknown"
    CALCULATING = "calculating"
    SUCCESS = "success"
    SKIPPED_LARGE = "skipped_large"
    SKIPPED_ENCRYPTED = "skipped_encrypted"
    SKIPPED_DAMAGED = "skipped_damaged"
    SKIPPED_NO_ACCESS = "skipped_no_access"
    SKIPPED_NO_OFFICE = "skipped_no_office"
    ERROR = "error"


@dataclass
class PageCountResult:
    """页数统计结果"""
    document: Document
    page_count: Optional[int] = None
    status: PageCountStatus = PageCountStatus.UNKNOWN
    error_message: str = ""
    calculation_time: float = 0.0


@dataclass
class PageCountSummary:
    """页数统计汇总"""
    total_files: int = 0
    total_pages: int = 0
    success_count: int = 0
    skipped_count: int = 0
    error_count: int = 0
    
    # 按文件类型统计
    word_files: int = 0
    word_pages: int = 0
    ppt_files: int = 0
    ppt_pages: int = 0
    excel_files: int = 0
    excel_pages: int = 0
    pdf_files: int = 0
    pdf_pages: int = 0
    image_files: int = 0
    image_pages: int = 0
    
    # 问题文件列表
    skipped_files: List[PageCountResult] = field(default_factory=list)
    error_files: List[PageCountResult] = field(default_factory=list)


class PageCountManager:
    """页数统计管理器 - 基于处理器架构"""

    def __init__(self, backend_preference: str = "auto", registry=None, backend_detector=None):
        self._cancel_flag = False
        self._progress_callback: Optional[Callable] = None
        self._backend_preference = backend_preference
        self._handler_registry = registry or create_default_registry()
        self._backend_detector = backend_detector or detect_backend
        self._pool: Optional[OfficeSessionPool] = None

    def set_progress_callback(self, callback: Callable[[int, int, str], None]):
        self._progress_callback = callback

    def cancel_calculation(self):
        self._cancel_flag = True

    def calculate_all_pages(self, documents: List[Document]) -> PageCountSummary:
        """批量计算所有文档的页数（顺序执行，调用方应在后台线程中调用）"""
        if not documents:
            return PageCountSummary()
        self._cancel_flag = False
        results: List[PageCountResult] = []
        total_docs = len(documents)
        needs_office = any(d.file_type in OFFICE_FILE_TYPES for d in documents)
        office_lock = (office_automation(5, "正在打印或检测 WPS，请稍后再统计 Office 文档页数")
                       if needs_office else nullcontext())
        with com_initialized(), office_lock:
            pool = None
            if needs_office:
                pool = OfficeSessionPool(self._backend_detector(self._backend_preference))
            ctx = PrintContext(office_pool=pool)
            try:
                for i, document in enumerate(documents):
                    if self._cancel_flag:
                        break
                    if self._progress_callback:
                        self._progress_callback(i + 1, total_docs, f"正在统计: {document.file_name}")
                    results.append(self._calculate_single_document(document, ctx, pool))
            finally:
                if pool is not None:
                    pool.close_all()

        summary = self._generate_summary(results)
        if self._progress_callback:
            self._progress_callback(len(results), total_docs, f"页数统计完成！总页数: {summary.total_pages}")
        logger.info("页数统计完成: %s/%s 成功", summary.success_count, len(results))
        return summary

    def _calculate_single_document(self, document: Document, ctx: Optional[PrintContext] = None,
                                   pool: Optional[OfficeSessionPool] = None) -> PageCountResult:
        start_time = time.time()
        result = PageCountResult(document=document)
        ctx = ctx or PrintContext()
        kind = KIND_BY_FILE_TYPE.get(document.file_type)

        def _on_timeout():
            if pool is not None and kind is not None:
                pool.kill_kind(kind, "页数统计超时")

        watchdog = Watchdog(PAGE_COUNT_TIMEOUT, _on_timeout, name=f"count-{document.file_name}")
        try:
            should_skip, skip_status = self._should_skip_document(document)
            if should_skip:
                result.status = skip_status
                result.error_message = self._get_skip_message(skip_status)
                return result
            handler = self._handler_registry.get_handler_by_file_type(document.file_type)
            if handler is None:
                result.status = PageCountStatus.ERROR
                result.error_message = f"不支持的文件类型: {document.file_type}"
                return result
            if not handler.can_handle_file(document.file_path):
                result.status = PageCountStatus.ERROR
                result.error_message = "处理器无法处理此文件"
                return result
            result.status = PageCountStatus.CALCULATING
            with watchdog:
                page_count = handler.count_pages(document.file_path, ctx)
            if watchdog.fired:
                raise RuntimeError(f"统计超时（>{PAGE_COUNT_TIMEOUT}s）")
            if page_count < 0:
                result.status = PageCountStatus.ERROR
                result.error_message = "获取到无效的页数"
            else:
                result.page_count = page_count
                result.status = PageCountStatus.SUCCESS
        except Exception as e:  # noqa: BLE001 - 单个文件失败不影响其它文件
            error_message = self._get_user_friendly_error(e, document.file_type)
            if "加密" in error_message or "密码" in error_message:
                result.status = PageCountStatus.SKIPPED_ENCRYPTED
            elif "损坏" in error_message:
                result.status = PageCountStatus.SKIPPED_DAMAGED
            elif "COM 接口不可用" in error_message or "未找到可用的 Office" in error_message \
                    or "需要安装" in error_message:
                result.status = PageCountStatus.SKIPPED_NO_OFFICE
            elif "无法访问" in error_message:
                result.status = PageCountStatus.SKIPPED_NO_ACCESS
            else:
                result.status = PageCountStatus.ERROR
            result.error_message = error_message
            logger.warning("页数统计失败: %s - %s", document.file_name, error_message)
        finally:
            result.calculation_time = time.time() - start_time
        return result

    def _should_skip_document(self, document: Document) -> Tuple[bool, PageCountStatus]:
        """
        检查是否应该跳过文档统计
        
        Args:
            document: 文档对象
            
        Returns:
            (是否跳过, 跳过状态)
        """
        # 检查文件是否存在
        if not document.file_path.exists():
            return True, PageCountStatus.SKIPPED_NO_ACCESS
        
        # 检查文件大小（超过100MB跳过）
        try:
            file_size_mb = document.file_path.stat().st_size / (1024 * 1024)
            if file_size_mb > 100:
                return True, PageCountStatus.SKIPPED_LARGE
        except OSError:
            pass
        
        return False, PageCountStatus.UNKNOWN
    
    def _get_user_friendly_error(self, error: Exception, file_type: FileType) -> str:
        """获取用户友好的错误信息"""
        error_str = str(error)
        if "文件被加密" in error_str or "密码保护" in error_str:
            return "文件被加密，无法统计页数"
        if "文件已损坏" in error_str:
            return "文件已损坏，无法读取"
        return f"统计页数失败: {error_str}"

    def _get_skip_message(self, status: PageCountStatus) -> str:
        """获取跳过原因的用户友好描述"""
        messages = {
            PageCountStatus.SKIPPED_LARGE: "文件过大（>100MB），已跳过",
            PageCountStatus.SKIPPED_ENCRYPTED: "文件被加密，无法访问",
            PageCountStatus.SKIPPED_DAMAGED: "文件已损坏，无法读取",
            PageCountStatus.SKIPPED_NO_ACCESS: "文件不存在或无法访问",
            PageCountStatus.SKIPPED_NO_OFFICE: "缺少必要的Office组件"
        }
        return messages.get(status, "未知原因跳过")
    
    def _generate_summary(self, results: List[PageCountResult]) -> PageCountSummary:
        """生成统计汇总"""
        summary = PageCountSummary()
        summary.total_files = len(results)
        
        for result in results:
            # 统计成功、跳过、错误数量
            if result.status == PageCountStatus.SUCCESS:
                summary.success_count += 1
                summary.total_pages += result.page_count or 0
                
                # 按文件类型统计
                if result.document.file_type == FileType.WORD:
                    summary.word_files += 1
                    summary.word_pages += result.page_count or 0
                elif result.document.file_type == FileType.PPT:
                    summary.ppt_files += 1
                    summary.ppt_pages += result.page_count or 0
                elif result.document.file_type == FileType.EXCEL:
                    summary.excel_files += 1
                    summary.excel_pages += result.page_count or 0
                elif result.document.file_type == FileType.PDF:
                    summary.pdf_files += 1
                    summary.pdf_pages += result.page_count or 0
                elif result.document.file_type == FileType.IMAGE:
                    summary.image_files += 1
                    summary.image_pages += result.page_count or 0
                    
            elif result.status in [
                PageCountStatus.SKIPPED_LARGE,
                PageCountStatus.SKIPPED_ENCRYPTED,
                PageCountStatus.SKIPPED_DAMAGED,
                PageCountStatus.SKIPPED_NO_ACCESS,
                PageCountStatus.SKIPPED_NO_OFFICE
            ]:
                summary.skipped_count += 1
                summary.skipped_files.append(result)
            else:
                summary.error_count += 1
                summary.error_files.append(result)
        
        return summary
    
    def get_supported_file_types(self) -> List[FileType]:
        return list(self._handler_registry.get_all_supported_file_types())

    def get_supported_extensions(self) -> List[str]:
        return sorted(self._handler_registry.get_all_supported_extensions())
