"""
批量打印控制器（WPS 版）

职责：
- 顺序执行打印队列，单个文件失败不会中断整个队列
- 为 Office 文档维护 WPS/Office 会话池（复用实例，批次结束统一退出并清理进程）
- 每个任务都有看门狗超时：WPS 卡死时强制结束本程序启动的 WPS 进程并继续下一个文件
- 批次期间临时设置打印机环境（每用户打印参数 + 必要时临时切换默认打印机），结束后恢复
- 为每个任务写入打印日志（logs/print_YYYYMMDD.log）
"""
import logging
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional

from ..backends.com_utils import describe_com_error
from ..backends.office_backend import (KIND_BY_FILE_TYPE, OfficeKind, OfficeSessionPool, Watchdog,
                                       com_initialized, detect_backend)
from ..handlers import create_default_registry
from ..utils.logging_setup import append_print_log, get_print_log_path
from .models import OFFICE_FILE_TYPES, Document, FileType, PrintSettings, PrintStatus
from .print_context import PrintCancelled, PrintContext, PrintJobError
from .printer_environment import PrinterEnvironment, PrinterEnvironmentBusy, get_default_printer

logger = logging.getLogger("bdp.print")

DEFAULT_TIMEOUTS = {
    OfficeKind.WRITER: 180,
    OfficeKind.SPREADSHEET: 300,
    OfficeKind.PRESENTATION: 300,
    None: 600,   # PDF/图片/文本（SumatraPDF 自身也有超时）
}


@dataclass
class TaskRecord:
    """单个打印任务的结果"""
    time: str
    file: str
    file_type: str
    backend: str
    printer: str
    paper: str
    duplex: str
    orientation: str
    copies: int
    color: str
    success: bool
    error: str = ""
    elapsed: float = 0.0
    notes: List[str] = field(default_factory=list)
    status: str = ""

    def to_log_line(self) -> str:
        fields = [
            self.time,
            self.status or ("成功" if self.success else "失败"),
            f"耗时={self.elapsed:.1f}s",
            f"类型={self.file_type}",
            f"Backend={self.backend or '-'}",
            f"打印机={self.printer}",
            f"纸张={self.paper}",
            f"单双面={self.duplex}",
            f"方向={self.orientation}",
            f"份数={self.copies}",
            f"色彩={self.color}",
            f"文件={self.file}",
        ]
        if self.error:
            fields.append(f"错误={self.error}")
        if self.notes:
            fields.append(f"备注={'; '.join(self.notes)}")
        return " | ".join(fields)


@dataclass
class BatchSummary:
    total: int = 0
    success: int = 0
    failed: int = 0
    cancelled: int = 0
    elapsed: float = 0.0
    records: List[TaskRecord] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    fatal_error: str = ""
    log_file: str = ""

    def text(self) -> str:
        parts = [f"成功 {self.success}", f"失败 {self.failed}"]
        if self.cancelled:
            parts.append(f"取消 {self.cancelled}")
        return "，".join(parts)


class PrintController:
    """批量打印控制器"""

    def __init__(self, registry=None, backend_detector=None, environment_factory=None,
                 pool_factory=None, timeouts=None, logs_dir: Optional[Path] = None,
                 inter_job_delay: float = 0.3):
        self._print_queue: List[Document] = []
        self._current_settings: Optional[PrintSettings] = None
        self._is_printing = False
        self._state_lock = threading.Lock()
        self._cancel_event = threading.Event()
        self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="PrintWorker")
        self._handler_registry = registry or create_default_registry()
        self._backend_detector = backend_detector or detect_backend
        self._environment_factory = environment_factory or self._default_environment
        self._pool_factory = pool_factory or (lambda selection, log: OfficeSessionPool(selection, log=log))
        self._timeouts = dict(DEFAULT_TIMEOUTS)
        if timeouts:
            self._timeouts.update(timeouts)
        self._logs_dir = logs_dir
        self._inter_job_delay = inter_job_delay
        self._active_pool: Optional[OfficeSessionPool] = None

        self._print_progress_callback: Optional[Callable[[int, int, str], None]] = None
        self._document_callback: Optional[Callable[[Document], None]] = None
        self._finished_callback: Optional[Callable[[BatchSummary], None]] = None
        self.last_summary: Optional[BatchSummary] = None

    @staticmethod
    def _default_environment(printer_name, settings, switch_default, log):
        return PrinterEnvironment(printer_name, settings, switch_default=switch_default, log=log)

    # ------------------------------------------------------------------ 配置
    def set_print_settings(self, settings: PrintSettings):
        self._current_settings = settings
        logger.info("打印设置已更新: 打印机=%s 纸张=%s", settings.printer_name, settings.paper_size)

    def set_progress_callback(self, callback: Callable[[int, int, str], None]):
        self._print_progress_callback = callback

    def set_document_callback(self, callback: Callable[[Document], None]):
        """单个文档状态变化时回调（在工作线程中调用）"""
        self._document_callback = callback

    def set_finished_callback(self, callback: Callable[[BatchSummary], None]):
        """批次结束时回调（在工作线程中调用）"""
        self._finished_callback = callback

    def add_documents_to_queue(self, documents: List[Document]):
        self._print_queue.extend(documents)

    def clear_queue(self):
        self._print_queue.clear()

    @property
    def queue_size(self) -> int:
        return len(self._print_queue)

    @property
    def is_printing(self) -> bool:
        return self._is_printing

    # ------------------------------------------------------------------ 执行
    def start_batch_print(self) -> Future:
        with self._state_lock:
            if self._is_printing:
                raise RuntimeError("打印任务已在进行中")
            if not self._print_queue:
                raise ValueError("打印队列为空")
            if not self._current_settings:
                raise ValueError("未设置打印参数")
            self._is_printing = True
            self._cancel_event.clear()
        documents = list(self._print_queue)
        settings = PrintSettings.from_dict(self._current_settings.to_dict())
        try:
            return self._executor.submit(self._execute_batch_print, documents, settings)
        except Exception:
            self._is_printing = False
            raise

    def cancel_current_print(self, force: bool = False):
        """
        请求取消：当前文件完成后停止。force=True 时立即结束本程序启动的 WPS 进程
        （用于退出程序），当前文件记为失败。
        """
        if not self._is_printing:
            return
        self._cancel_event.set()
        logger.info("请求取消打印（force=%s）", force)
        pool = self._active_pool
        if force and pool is not None:
            for kind in OfficeKind:
                pool.kill_kind(kind, "用户取消/退出程序")

    def _log(self, message: str) -> None:
        logger.info(message)

    def _notify_document(self, document: Document) -> None:
        if self._document_callback:
            try:
                self._document_callback(document)
            except Exception:  # noqa: BLE001 - 回调异常不能影响打印
                logger.exception("文档状态回调失败")

    def _notify_progress(self, current: int, total: int, message: str) -> None:
        if self._print_progress_callback:
            try:
                self._print_progress_callback(current, total, message)
            except Exception:  # noqa: BLE001
                logger.exception("进度回调失败")

    def _execute_batch_print(self, documents: List[Document], settings: PrintSettings) -> BatchSummary:
        summary = BatchSummary(total=len(documents))
        summary.log_file = str(get_print_log_path(logs_dir=self._logs_dir))
        batch_start = time.monotonic()
        for doc in documents:
            doc.print_status = PrintStatus.PENDING
            doc.last_error = ""
            self._notify_document(doc)
        try:
            if not settings.printer_name:
                settings.printer_name = get_default_printer() or ""
            has_office = any(doc.file_type in OFFICE_FILE_TYPES for doc in documents)
            ctx = PrintContext(log=self._log, cancel_event=self._cancel_event)
            header = (f"===== 批次开始 {datetime.now():%Y-%m-%d %H:%M:%S} 共 {len(documents)} 个文件 | "
                      f"打印机={settings.printer_name or '(无)'} | 纸张={settings.paper_size} | "
                      f"{settings.duplex_display} | 方向={settings.orientation_display} | "
                      f"份数={settings.copies} | {settings.color_display} | "
                      f"Office引擎偏好={settings.office_backend.value}")
            append_print_log(header, self._logs_dir)
            logger.info(header)

            with com_initialized():
                pool = None
                if has_office:
                    selection = self._backend_detector(settings.office_backend.value)
                    pool = self._pool_factory(selection, self._log)
                    self._active_pool = pool
                    for kind in OfficeKind:
                        backend, message = selection.backend_for(kind)
                        self._log(f"Office 组件 {kind.value}: {message.splitlines()[0] if message else ''}")
                    ctx.office_pool = pool
                try:
                    with self._environment_factory(settings.printer_name, settings, has_office, self._log) as env:
                        ctx.default_printer_switched = bool(getattr(env, "default_switched", False))
                        summary.warnings = list(getattr(env, "warnings", []) or [])
                        for warning in summary.warnings:
                            append_print_log(f"  ⚠ {warning}", self._logs_dir)
                        try:
                            self._run_queue(documents, settings, ctx, pool, summary)
                        finally:
                            if pool is not None:
                                pool.close_all()
                finally:
                    self._active_pool = None
        except PrinterEnvironmentBusy as e:
            summary.fatal_error = str(e)
        except Exception as e:  # noqa: BLE001 - 批次级异常：记录并结束，不让线程崩溃
            logger.exception("批量打印发生异常")
            summary.fatal_error = f"批量打印发生异常: {describe_com_error(e)}"
        finally:
            if summary.fatal_error:
                for doc in documents:
                    if doc.print_status in (PrintStatus.PENDING, PrintStatus.PRINTING):
                        doc.print_status = PrintStatus.ERROR
                        doc.last_error = summary.fatal_error
                        summary.failed += 1
                        self._notify_document(doc)
                append_print_log(f"  ✖ {summary.fatal_error}", self._logs_dir)
            summary.elapsed = time.monotonic() - batch_start
            footer = f"===== 批次结束 {summary.text()} 共 {summary.total} 个，耗时 {summary.elapsed:.1f}s"
            append_print_log(footer, self._logs_dir)
            logger.info(footer)
            self.last_summary = summary
            self._is_printing = False
            self._notify_progress(summary.total, summary.total, f"批量打印完成！{summary.text()}")
            if self._finished_callback:
                try:
                    self._finished_callback(summary)
                except Exception:  # noqa: BLE001
                    logger.exception("完成回调失败")
        return summary

    def _run_queue(self, documents: List[Document], settings: PrintSettings, ctx: PrintContext,
                   pool: Optional[OfficeSessionPool], summary: BatchSummary) -> None:
        total = len(documents)
        for index, document in enumerate(documents):
            if self._cancel_event.is_set():
                for rest in documents[index:]:
                    rest.print_status = PrintStatus.CANCELLED
                    rest.last_error = "用户取消"
                    summary.cancelled += 1
                    self._notify_document(rest)
                    append_print_log(self._make_record(rest, settings, "", False, "用户取消", 0.0, [],
                                                       status="已取消").to_log_line(), self._logs_dir)
                break
            self._notify_progress(index + 1, total, f"正在打印: {document.file_name}")
            record = self._print_one(document, settings, ctx, pool)
            summary.records.append(record)
            if record.success:
                summary.success += 1
            else:
                summary.failed += 1
            if self._inter_job_delay:
                time.sleep(self._inter_job_delay)

    def _make_record(self, document: Document, settings: PrintSettings, backend: str, success: bool,
                     error: str, elapsed: float, notes: List[str], status: str = "") -> TaskRecord:
        return TaskRecord(
            time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            file=str(document.file_path),
            file_type=f"{document.type_display}({document.file_path.suffix.lower()})",
            backend=backend,
            printer=settings.printer_name,
            paper=settings.paper_size,
            duplex=settings.duplex_display,
            orientation=settings.orientation_display,
            copies=settings.copies,
            color=settings.color_display,
            success=success,
            error=error,
            elapsed=elapsed,
            notes=list(notes),
            status=status,
        )

    def _print_one(self, document: Document, settings: PrintSettings, ctx: PrintContext,
                   pool: Optional[OfficeSessionPool]) -> TaskRecord:
        document.print_status = PrintStatus.PRINTING
        document.last_error = ""
        self._notify_document(document)
        ctx.reset_job()
        kind = KIND_BY_FILE_TYPE.get(document.file_type)
        timeout = self._timeouts.get(kind, self._timeouts[None])
        handler = self._handler_registry.get_handler_by_file_type(document.file_type)
        start = time.monotonic()
        error = ""

        def _on_timeout():
            if kind is not None and pool is not None:
                pool.kill_kind(kind, f"打印超时（>{timeout}s）")

        watchdog = Watchdog(timeout, _on_timeout, name=document.file_name)
        try:
            if not document.file_path.exists():
                raise PrintJobError("文件不存在或无法访问（可能已被移动/删除，或网络共享已断开）")
            if handler is None:
                raise PrintJobError(f"不支持的文件类型: {document.file_path.suffix}")
            with watchdog:
                ok = handler.print_document(document.file_path, settings, ctx)
            if watchdog.fired:
                raise PrintJobError("打印超时")
            if not ok:
                raise PrintJobError("处理器返回失败（详见 logs/app.log）")
        except PrintCancelled:
            error = "用户取消"
        except PrintJobError as e:
            error = str(e)
        except Exception as e:  # noqa: BLE001 - 任何异常都只影响当前文件
            logger.exception("打印 %s 时发生未预期的异常", document.file_path)
            error = describe_com_error(e)
        if watchdog.fired:
            component = kind.value if kind else "打印程序"
            error = (f"操作超时（超过 {timeout} 秒无响应），已强制结束本程序启动的 {component} 进程"
                     + (f"；{error}" if error and error != "打印超时" else ""))

        elapsed = time.monotonic() - start
        success = not error
        backend = handler.describe_backend(ctx) if handler is not None else ""
        document.last_backend = backend
        document.print_status = PrintStatus.COMPLETED if success else PrintStatus.ERROR
        document.last_error = error
        record = self._make_record(document, settings, backend, success, error, elapsed, ctx.notes)
        append_print_log(record.to_log_line(), self._logs_dir)
        (logger.info if success else logger.error)(record.to_log_line())
        self._notify_document(document)
        return record

    # ------------------------------------------------------------------ 查询
    def get_print_queue_status(self) -> dict:
        status_count = {status.value: 0 for status in PrintStatus}
        for doc in self._print_queue:
            status_count[doc.print_status.value] += 1
        return {
            'total': len(self._print_queue),
            'status_count': status_count,
            'is_printing': self._is_printing,
            'supported_types': [ft.value for ft in self._handler_registry.get_all_supported_file_types()],
            'supported_extensions': sorted(self._handler_registry.get_all_supported_extensions())
        }

    def get_supported_file_types(self) -> List[FileType]:
        return list(self._handler_registry.get_all_supported_file_types())

    def get_supported_extensions(self) -> List[str]:
        return sorted(self._handler_registry.get_all_supported_extensions())

    def shutdown(self, wait: bool = True):
        self._executor.shutdown(wait=wait)
