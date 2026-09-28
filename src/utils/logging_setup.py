"""
日志配置

- logs/app.log            程序运行日志（滚动，5MB x 5）
- logs/print_YYYYMMDD.log 打印任务日志（每天一个文件，每个任务一行）

程序中遗留的 print() 输出通过 StreamToLogger 转发到 app.log，
因此在无控制台的 exe 中也能留下诊断信息。
"""
import logging
import sys
import threading
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

from src.utils.path_utils import get_logs_dir

APP_LOGGER_NAME = "bdp"
_LOG_FORMAT = "%(asctime)s [%(levelname)s] %(threadName)s %(name)s: %(message)s"
_configured = False
_print_log_lock = threading.Lock()


class StreamToLogger:
    """把写入 stdout/stderr 的内容转发给 logger，同时保留原始流（如果存在）"""

    def __init__(self, logger: logging.Logger, level: int, original=None):
        self._logger = logger
        self._level = level
        self._original = original
        self._buffer = ""
        self._lock = threading.Lock()

    def write(self, message):
        if not message:
            return 0
        if self._original is not None:
            try:
                self._original.write(message)
            except Exception:  # 原始控制台不可用（例如编码问题）时忽略
                pass
        with self._lock:
            self._buffer += message
            while "\n" in self._buffer:
                line, self._buffer = self._buffer.split("\n", 1)
                line = line.rstrip()
                if line:
                    self._logger.log(self._level, line)
        return len(message)

    def flush(self):
        if self._original is not None:
            try:
                self._original.flush()
            except Exception:
                pass

    def isatty(self):
        return False

    @property
    def encoding(self):
        return "utf-8"


def setup_logging(level: int = logging.INFO, redirect_std: bool = True,
                  logs_dir: Optional[Path] = None) -> Path:
    """
    初始化日志系统（可重复调用，只生效一次）

    Returns:
        日志目录
    """
    global _configured
    logs_dir = logs_dir or get_logs_dir()
    if _configured:
        return logs_dir

    logger = logging.getLogger(APP_LOGGER_NAME)
    logger.setLevel(level)
    logger.propagate = False

    file_handler = RotatingFileHandler(
        logs_dir / "app.log", maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(logging.Formatter(_LOG_FORMAT))
    logger.addHandler(file_handler)

    if redirect_std:
        stdout_logger = logging.getLogger(f"{APP_LOGGER_NAME}.stdout")
        stderr_logger = logging.getLogger(f"{APP_LOGGER_NAME}.stderr")
        sys.stdout = StreamToLogger(stdout_logger, logging.INFO, sys.stdout)
        sys.stderr = StreamToLogger(stderr_logger, logging.ERROR, sys.stderr)

        def _excepthook(exc_type, exc, tb):
            logger.critical("未捕获的异常", exc_info=(exc_type, exc, tb))

        sys.excepthook = _excepthook

        def _thread_excepthook(args):
            logger.critical("线程 %s 未捕获的异常", getattr(args.thread, "name", "?"),
                            exc_info=(args.exc_type, args.exc_value, args.exc_traceback))

        threading.excepthook = _thread_excepthook

    _configured = True
    logger.info("日志系统已初始化，目录: %s", logs_dir)
    return logs_dir


def get_logger(name: str) -> logging.Logger:
    """获取应用子 logger"""
    if name.startswith(APP_LOGGER_NAME):
        return logging.getLogger(name)
    return logging.getLogger(f"{APP_LOGGER_NAME}.{name}")


def get_print_log_path(when: Optional[datetime] = None, logs_dir: Optional[Path] = None) -> Path:
    when = when or datetime.now()
    logs_dir = logs_dir or get_logs_dir()
    return logs_dir / f"print_{when.strftime('%Y%m%d')}.log"


def append_print_log(line: str, logs_dir: Optional[Path] = None) -> None:
    """向当天的打印日志追加一行（线程安全）"""
    path = get_print_log_path(logs_dir=logs_dir)
    with _print_log_lock:
        try:
            with open(path, "a", encoding="utf-8") as f:
                f.write(line.rstrip("\n") + "\n")
        except OSError as e:
            logging.getLogger(APP_LOGGER_NAME).error("写入打印日志失败 %s: %s", path, e)
