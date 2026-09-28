"""
打印上下文：批量打印期间在控制器与处理器之间传递的共享对象
"""
import logging
import threading
from dataclasses import dataclass, field
from typing import Callable, List, Optional

logger = logging.getLogger("bdp.print")


class PrintJobError(Exception):
    """单个打印任务失败（message 为面向用户的中文说明）"""


class PrintCancelled(Exception):
    """用户取消了打印"""


@dataclass
class PrintContext:
    office_pool: Optional[object] = None           # OfficeSessionPool
    log: Callable[[str], None] = field(default=lambda msg: logger.info(msg))
    cancel_event: threading.Event = field(default_factory=threading.Event)
    default_printer_switched: bool = False
    # 以下为单个任务级别的信息，每个任务开始前由控制器重置
    backend_used: str = ""
    notes: List[str] = field(default_factory=list)

    def reset_job(self) -> None:
        self.backend_used = ""
        self.notes = []

    def check_cancelled(self) -> None:
        if self.cancel_event.is_set():
            raise PrintCancelled("用户取消了打印")
