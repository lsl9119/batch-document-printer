"""
Batch Document Printer – WPS Edition 程序入口

Windows + WPS Office 批量打印工具（基于 batch-document-printer 二次开发）
"""
import os
import sys
from pathlib import Path

# 保证以 "src.xxx" 方式导入（开发环境与 PyInstaller 打包环境一致）
_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from src.version import APP_NAME, VERSION  # noqa: E402


def _show_fatal(message: str) -> None:
    try:
        import tkinter.messagebox as messagebox
        messagebox.showerror(APP_NAME, message)
    except Exception:  # noqa: BLE001 - 无法显示对话框时只写 stderr
        print(message, file=sys.stderr)


def check_system_requirements() -> None:
    """检查系统要求"""
    import platform
    if platform.system() != "Windows":
        raise RuntimeError("此应用程序仅支持 Windows 操作系统")
    if sys.version_info < (3, 8):
        raise RuntimeError("需要 Python 3.8 或更高版本")
    missing = []
    for module in ("tkinter", "win32print", "win32api", "win32com.client", "pythoncom"):
        try:
            __import__(module)
        except ImportError:
            missing.append(module)
    if missing:
        raise RuntimeError(f"缺少必要的依赖模块: {', '.join(missing)}\n请运行: pip install -r requirements.txt")


def startup_maintenance(logger) -> None:
    """启动时的维护工作：恢复上次异常退出遗留的打印机状态、清理临时文件"""
    from src.core.printer_environment import recover_pending_state
    from src.core.file_prep import cleanup_stale_temp
    try:
        for action in recover_pending_state(log=logger.warning):
            logger.warning(action)
    except Exception:  # noqa: BLE001 - 恢复失败不能阻止程序启动
        logger.exception("恢复打印机状态失败")
    try:
        removed = cleanup_stale_temp()
        if removed:
            logger.info("已清理 %d 个残留临时目录", removed)
    except OSError:
        logger.exception("清理临时目录失败")


def main() -> int:
    from src.utils.logging_setup import setup_logging, get_logger
    setup_logging()
    logger = get_logger("main")
    logger.info("=" * 60)
    logger.info("%s v%s 启动 (pid=%s, frozen=%s)", APP_NAME, VERSION, os.getpid(),
                getattr(sys, "frozen", False))
    try:
        check_system_requirements()
        startup_maintenance(logger)
        from src.gui.main_window import MainWindow
        # 拖到 exe 图标上的文件/文件夹会作为命令行参数传入
        initial_paths = [arg for arg in sys.argv[1:] if not arg.startswith("-")]
        app = MainWindow(initial_paths=initial_paths)
        app.run()
        return 0
    except Exception as e:  # noqa: BLE001 - 顶层兜底，记录并提示
        logger.exception("应用程序运行失败")
        _show_fatal(f"应用程序运行失败:\n{e}\n\n详细信息见日志 logs/app.log")
        return 1


if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        import multiprocessing
        multiprocessing.freeze_support()
    sys.exit(main())
