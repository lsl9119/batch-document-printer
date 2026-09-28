"""
SumatraPDF 命令行打印封装（PDF 与图片共用）

文档: https://www.sumatrapdfreader.org/docs/Command-line-arguments
-print-settings 支持: 份数 Nx、simplex/duplexlong/duplexshort、portrait/landscape、
                     color/monochrome、fit/shrink/noscale、paper=A4 等
"""
import logging
import os
import subprocess
import tempfile
import threading
from pathlib import Path
from typing import List, Optional, Set

from ..core.models import Orientation, PrintSettings
from ..core.print_context import PrintJobError
from ..utils.path_utils import get_sumatra_pdf_path

logger = logging.getLogger("bdp.sumatra")

# SumatraPDF 认识的纸张名称
SUMATRA_PAPERS = {"a2": "A2", "a3": "A3", "a4": "A4", "a5": "A5", "a6": "A6",
                  "letter": "letter", "legal": "legal", "tabloid": "tabloid", "statement": "statement"}

CREATE_NO_WINDOW = 0x08000000

# 正在运行的外部打印进程（用于取消/看门狗强制结束）
_active: Set[subprocess.Popen] = set()
_active_lock = threading.Lock()


class ExternalPrintTimeout(PrintJobError):
    """外部打印程序超时：作业可能已部分进入打印队列，不应自动换用其他方式重打"""


def kill_active(reason: str = "") -> None:
    with _active_lock:
        procs = list(_active)
    for proc in procs:
        _kill_tree(proc, reason or "强制结束")


def sumatra_path() -> Path:
    return get_sumatra_pdf_path()


def sumatra_available() -> bool:
    return sumatra_path().is_file()


def build_print_settings(settings: PrintSettings) -> List[str]:
    parts = [f"{max(1, int(settings.copies))}x", settings.duplex_mode_str]
    if settings.orientation != Orientation.AUTO:
        parts.append(settings.orientation.value)
    parts.append("color" if settings.color else "monochrome")
    parts.append(settings.scaling.value)
    paper = SUMATRA_PAPERS.get((settings.paper_size or "").strip().lower())
    if paper:
        parts.append(f"paper={paper}")
    return parts


def build_command(file_path: Path, printer_name: str, settings: PrintSettings,
                  exe: Optional[Path] = None) -> List[str]:
    exe = exe or sumatra_path()
    return [str(exe), "-print-to", printer_name,
            "-print-settings", ",".join(build_print_settings(settings)),
            "-silent", str(file_path)]


def compute_timeout(file_path: Path, base: int = 120) -> int:
    try:
        size_mb = file_path.stat().st_size / (1024 * 1024)
    except OSError:
        size_mb = 0
    return int(min(900, base + size_mb * 10))


def run_print(file_path: Path, printer_name: str, settings: PrintSettings, log=None) -> str:
    """执行打印；失败抛出 PrintJobError。返回实际使用的 print-settings 字符串"""
    log = log or (lambda m: logger.info(m))
    if not sumatra_available():
        raise PrintJobError(f"SumatraPDF 不存在: {sumatra_path()}")
    if not printer_name:
        raise PrintJobError("未指定打印机")
    cmd = build_command(file_path, printer_name, settings)
    timeout = compute_timeout(file_path)
    log(f"SumatraPDF 打印: -print-to \"{printer_name}\" -print-settings {cmd[4]} (超时 {timeout}s)")
    try:
        returncode, detail = run_detached(cmd, timeout)
    except subprocess.TimeoutExpired as e:
        raise ExternalPrintTimeout(f"SumatraPDF 打印超时（{timeout} 秒），已结束 SumatraPDF 进程。"
                                   f"作业可能已部分发送，请检查打印机队列后再决定是否重打") from e
    except OSError as e:
        raise PrintJobError(f"无法启动 SumatraPDF: {e}") from e
    if returncode != 0:
        raise PrintJobError(f"SumatraPDF 打印失败（退出码 {returncode}）: {detail[:300] or '无输出'}")
    return cmd[4]


def run_detached(cmd: List[str], timeout: float):
    """
    运行外部程序并等待其退出。

    不使用管道捕获输出：若外部程序的子进程继承了管道句柄，subprocess.run(capture_output=True)
    会在程序退出后继续等待管道关闭而无限阻塞（超时参数对此无效）。
    这里把 stderr 写入临时文件，只等待进程本身；超时则结束整个进程树。

    Returns:
        (退出码, stderr 文本)
    """
    kwargs = {}
    if os.name == "nt":
        kwargs["creationflags"] = CREATE_NO_WINDOW
    with tempfile.TemporaryFile() as err:
        proc = subprocess.Popen(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=err, **kwargs)
        with _active_lock:
            _active.add(proc)
        try:
            returncode = proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            _kill_tree(proc, "外部打印程序超时")
            raise
        finally:
            with _active_lock:
                _active.discard(proc)
        err.seek(0)
        detail = err.read().decode("utf-8", errors="replace").strip()
    return returncode, detail


def _kill_tree(proc: subprocess.Popen, reason: str) -> None:
    from ..backends import process_utils
    try:
        import psutil  # type: ignore
        children = [c.pid for c in psutil.Process(proc.pid).children(recursive=True)]
    except Exception:  # noqa: BLE001 - 取不到子进程时只结束主进程
        children = []
    process_utils.kill_pids(children + [proc.pid], reason=reason)
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        logger.error("外部进程 pid=%s 无法结束", proc.pid)
