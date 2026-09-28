"""
环境诊断与关于对话框
"""
import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox

from src.core.environment import EnvironmentReport, collect_environment
from src.utils.path_utils import get_logs_dir
from src.version import APP_NAME, VERSION, UPSTREAM_PROJECT, UPSTREAM_AUTHOR, PROJECT_URL


def _text_window(parent, title: str, geometry: str = "720x600"):
    window = tk.Toplevel(parent)
    window.title(title)
    window.geometry(geometry)
    window.transient(parent)
    frame = ttk.Frame(window, padding=10)
    frame.pack(fill="both", expand=True)
    text = tk.Text(frame, wrap="word", font=("Microsoft YaHei", 10), relief="flat")
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    buttons = ttk.Frame(window, padding=(10, 0, 10, 10))
    buttons.pack(fill="x", side="bottom")
    text.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    return window, text, buttons


def _set_text(widget: tk.Text, content: str) -> None:
    widget.config(state="normal")
    widget.delete("1.0", "end")
    widget.insert("1.0", content)
    widget.config(state="disabled")


def open_logs_folder() -> None:
    try:
        os.startfile(str(get_logs_dir()))  # type: ignore[attr-defined]
    except (AttributeError, OSError) as e:
        messagebox.showerror("错误", f"无法打开日志目录: {e}")


def show_environment_dialog(parent, report: EnvironmentReport, preference: str = "auto",
                            on_updated=None) -> None:
    window, text, buttons = _text_window(parent, "环境诊断")
    _set_text(text, report.text())

    def run_deep_check():
        deep_btn.config(state="disabled")
        _set_text(text, "正在启动 WPS 文字 / WPS 表格 / WPS 演示 进行实际 COM 测试，请稍候（约 10~60 秒）...\n"
                        "测试会创建隐藏的 WPS 进程并立即退出，不会打印任何内容。")
        results: "queue.Queue" = queue.Queue()

        def worker():
            try:
                new_report = collect_environment(preference, deep=True)
                results.put((new_report, new_report.text()))
            except Exception as e:  # noqa: BLE001 - 诊断失败也要展示给用户（包括“正在打印”）
                results.put((None, f"深度检测未执行/失败: {e}"))

        def poll():
            # 只在主线程中操作 Tkinter 控件
            try:
                new_report, content = results.get_nowait()
            except queue.Empty:
                parent.after(200, poll)
                return
            if not window.winfo_exists():
                return
            _set_text(text, content)
            deep_btn.config(state="normal")
            if new_report is not None and on_updated:
                on_updated(new_report)

        # 非守护线程：程序退出时会等待检测结束，确保测试启动的 WPS 进程被正确退出
        threading.Thread(target=worker, name="DeepEnvCheck", daemon=False).start()
        parent.after(200, poll)

    def copy_report():
        window.clipboard_clear()
        window.clipboard_append(text.get("1.0", "end"))

    deep_btn = ttk.Button(buttons, text="深度检测（实际启动 COM）", command=run_deep_check)
    deep_btn.pack(side="left")
    ttk.Button(buttons, text="复制报告", command=copy_report).pack(side="left", padx=5)
    ttk.Button(buttons, text="打开日志目录", command=open_logs_folder).pack(side="left")
    ttk.Button(buttons, text="关闭", command=window.destroy).pack(side="right")


ABOUT_TEXT = f"""{APP_NAME}
版本 v{VERSION}

Windows + WPS Office 批量打印工具：只安装 WPS Office、不安装 Microsoft Office，
即可批量打印 PDF、DOC/DOCX/WPS、XLS/XLSX/ET、PPT/PPTX/DPS、图片与 TXT。

项目地址: {PROJECT_URL}

致谢与许可
• 基于 batch-document-printer（办公文档批量打印器，作者 {UPSTREAM_AUTHOR}，MIT 许可）
  {UPSTREAM_PROJECT}
• WPS COM 调用方式参考 harness-anything（cli-anything-wps contributors，MIT 许可）
  https://github.com/yb2460/harness-anything
• PDF/图片打印使用 SumatraPDF（GPLv3，作为独立程序随附）
  https://github.com/sumatrapdfreader/sumatrapdf
• 详见程序目录中的 LICENSE 与 NOTICE 文件
"""


def show_about_dialog(parent) -> None:
    window, text, buttons = _text_window(parent, "关于", "620x420")
    _set_text(text, ABOUT_TEXT)
    ttk.Button(buttons, text="关闭", command=window.destroy).pack(side="right")
