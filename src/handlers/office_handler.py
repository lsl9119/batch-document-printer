"""
Office 类文档处理器公共基类（文字 / 表格 / 演示）

处理器不直接创建 COM 对象，而是通过 OfficeSessionPool 获取对应组件的会话：
- 批量打印时由 PrintController 提供会话池（复用 WPS 实例，批次结束统一退出）
- 单独调用时（如页数统计、脚本）自动创建临时会话池并在结束时清理
"""
from contextlib import contextmanager
from pathlib import Path
from typing import Optional, Set

from ..backends.com_utils import describe_com_error, is_dead_server_error
from ..backends.office_backend import (BackendUnavailableError, OfficeKind, OfficeSession,
                                       OfficeSessionPool, OfficeStartError, com_initialized,
                                       detect_backend)
from ..backends.office_ops import OfficeDocumentError
from ..core.file_prep import PreparedFile
from ..core.models import FileType, PrintSettings
from ..core.print_context import PrintContext, PrintJobError
from .base_handler import BaseDocumentHandler


class OfficeDocumentHandler(BaseDocumentHandler):
    """Office 文档处理器基类"""

    requires_office = True
    KIND: OfficeKind = OfficeKind.WRITER
    FILE_TYPE: FileType = FileType.WORD
    EXTENSIONS: Set[str] = set()

    def __init__(self, backend_preference: str = "auto"):
        self.backend_preference = backend_preference

    def get_supported_file_types(self) -> Set[FileType]:
        return {self.FILE_TYPE}

    def get_supported_extensions(self) -> Set[str]:
        return set(self.EXTENSIONS)

    def describe_backend(self, context: Optional[PrintContext] = None) -> str:
        if context is not None and context.backend_used:
            return context.backend_used
        return self.get_handler_name()

    # ------------------------------------------------------------------
    @contextmanager
    def _pool_scope(self, context: Optional[PrintContext], settings: Optional[PrintSettings]):
        """优先使用上下文中的会话池；没有则创建临时会话池并在结束时关闭"""
        if context is not None and context.office_pool is not None:
            yield context.office_pool, context
            return
        context = context or PrintContext()
        preference = settings.office_backend.value if settings is not None else self.backend_preference
        with com_initialized():
            pool = OfficeSessionPool(detect_backend(preference), log=context.log)
            try:
                yield pool, context
            finally:
                pool.close_all()

    def _run(self, file_path: Path, settings: Optional[PrintSettings], context: Optional[PrintContext],
             operation):
        with self._pool_scope(context, settings) as (pool, ctx):
            try:
                session = pool.get(self.KIND)
            except (BackendUnavailableError, OfficeStartError) as e:
                raise PrintJobError(str(e)) from e
            ctx.backend_used = f"{session.backend.display_name} ({session.progid})"
            try:
                with PreparedFile(file_path, log=ctx.log) as prepared:
                    result = operation(session, prepared.path, ctx)
                session.jobs_done += 1
                return result
            except PrintJobError:
                self._discard_if_broken(pool, session)
                raise
            except OfficeDocumentError as e:
                self._discard_if_broken(pool, session)
                raise PrintJobError(str(e)) from e
            except FileNotFoundError as e:
                raise PrintJobError(str(e)) from e
            except Exception as e:  # noqa: BLE001 - COM 异常统一转换为可读信息
                broken = self._discard_if_broken(pool, session, force=is_dead_server_error(e))
                message = describe_com_error(e)
                if broken:
                    message += f"（{session.backend.component_name(self.KIND)} 已失去响应，已重启该组件）"
                raise PrintJobError(message) from e

    def _discard_if_broken(self, pool: OfficeSessionPool, session: OfficeSession, force: bool = False) -> bool:
        if force or not session.is_alive:
            if pool.peek(self.KIND) is session:
                pool.discard(self.KIND)
            return True
        return False

    # ------------------------------------------------------------------
    def print_document(self, file_path: Path, settings: PrintSettings,
                       context: Optional[PrintContext] = None) -> bool:
        def _op(session: OfficeSession, path: Path, ctx: PrintContext):
            return self._print_with_session(session, path, settings, ctx)

        self._run(file_path, settings, context, _op)
        return True

    def count_pages(self, file_path: Path, context: Optional[PrintContext] = None) -> int:
        def _op(session: OfficeSession, path: Path, ctx: PrintContext):
            return self._count_with_session(session, path, ctx)

        return int(self._run(file_path, None, context, _op))

    # 子类实现 --------------------------------------------------------
    def _print_with_session(self, session: OfficeSession, path: Path, settings: PrintSettings,
                            ctx: PrintContext) -> None:
        raise NotImplementedError

    def _count_with_session(self, session: OfficeSession, path: Path, ctx: PrintContext) -> int:
        raise NotImplementedError
