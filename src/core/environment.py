"""
运行环境检测

启动时（后台线程）检测：Windows 版本、WPS 文字/表格/演示 COM、Microsoft Office COM（可选回退）、
SumatraPDF、打印机列表与默认打印机。结果写入日志并显示在主界面。

WPS 检测依据 COM 注册（ProgID → CLSID → LocalServer32 可执行文件），而不是安装目录；
"深度检测"会真正创建 COM 对象并退出，用于诊断注册正常但无法启动的情况。
"""
import logging
import platform
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from ..backends.office_backend import (BackendSelection, ComponentStatus, OfficeKind, com_initialized,
                                       default_backends, office_automation, select_backends)
from ..backends import process_utils
from ..utils import path_utils
from ..version import APP_NAME, VERSION

logger = logging.getLogger("bdp.env")


def windows_version() -> str:
    if sys.platform != "win32":
        return f"{platform.system()} {platform.release()}（非 Windows，无法打印）"
    release, version, csd, _ = platform.win32_ver()
    try:
        build = int(version.split(".")[2])
    except (IndexError, ValueError):
        build = 0
    name = "Windows 11" if build >= 22000 else f"Windows {release}"
    edition = ""
    try:
        edition = platform.win32_edition() or ""
    except AttributeError:
        pass
    return f"{name} {edition} (版本 {version}{', ' + csd if csd else ''}, {platform.machine()})".replace("  ", " ")


def list_printers() -> List[str]:
    try:
        import win32print  # type: ignore
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return [p[2] for p in win32print.EnumPrinters(flags)]
    except Exception as e:  # noqa: BLE001
        logger.warning("枚举打印机失败: %s", e)
        return []


@dataclass
class EnvironmentReport:
    app: str = f"{APP_NAME} v{VERSION}"
    windows: str = ""
    python: str = ""
    selection: Optional[BackendSelection] = None
    statuses: Dict[str, Dict[OfficeKind, ComponentStatus]] = field(default_factory=dict)
    sumatra_path: str = ""
    sumatra_ok: bool = False
    printers: List[str] = field(default_factory=list)
    default_printer: str = ""
    running_office: List[str] = field(default_factory=list)
    data_dir: str = ""
    logs_dir: str = ""
    deep: bool = False

    @property
    def wps(self) -> Dict[OfficeKind, ComponentStatus]:
        return self.statuses.get("wps", {})

    @property
    def msoffice(self) -> Dict[OfficeKind, ComponentStatus]:
        return self.statuses.get("msoffice", {})

    @property
    def wps_detected(self) -> bool:
        return any(s.registered for s in self.wps.values())

    @property
    def wps_complete(self) -> bool:
        return bool(self.wps) and all(s.available for s in self.wps.values())

    def backend_label(self) -> str:
        if self.wps_complete:
            return "WPS Office ✓"
        if self.selection is not None and self.selection.any_available:
            return f"{self.selection.primary_display} ⚠" if self.wps_detected else f"{self.selection.primary_display} ✓"
        return "未检测到 WPS/Office ✗"

    def problems(self) -> List[str]:
        problems: List[str] = []
        if self.wps_detected:
            for status in self.wps.values():
                if not status.available:
                    problems.append(f"检测到 WPS Office，但{status.display_name} COM 接口不可用。"
                                    f"ProgID: {status.progid}" + (f"（{status.error}）" if status.error else ""))
        if self.selection is not None:
            for kind in OfficeKind:
                backend, message = self.selection.backend_for(kind)
                if backend is None and not self.wps_detected:
                    problems.append(message.splitlines()[0])
                    break
        if not self.sumatra_ok:
            problems.append(f"未找到 SumatraPDF（{self.sumatra_path}），PDF 无法打印，图片将使用 GDI 备用方案")
        if not self.printers:
            problems.append("未检测到任何打印机")
        if self.running_office:
            problems.append("检测到 WPS/Office 正在运行：" + ", ".join(self.running_office) +
                            "。建议打印前保存并关闭，以免批量打印附着到已打开的实例")
        # 去重保持顺序
        seen, unique = set(), []
        for p in problems:
            if p not in seen:
                seen.add(p)
                unique.append(p)
        return unique

    def text(self) -> str:
        lines = [self.app, f"操作系统: {self.windows}", f"Python: {self.python}", ""]
        lines.append(f"Office Backend: {self.backend_label()}")
        if self.selection is not None:
            for kind in OfficeKind:
                backend, _ = self.selection.backend_for(kind)
                lines.append(f"  {kind.value:<12} → {backend.display_name if backend else '不可用'}")
        lines.append("")
        lines.append("WPS Office COM（{}）:".format("深度检测" if self.deep else "注册表检测"))
        for status in self.wps.values():
            lines.extend("  " + l for l in status.detail_lines())
        lines.append("")
        lines.append("Microsoft Office COM（可选回退）:")
        for status in self.msoffice.values():
            lines.extend("  " + l for l in status.detail_lines())
        lines.append("")
        lines.append(f"SumatraPDF: {'✓' if self.sumatra_ok else '✗'} {self.sumatra_path}")
        lines.append(f"默认打印机: {self.default_printer or '(无)'}")
        lines.append(f"打印机 ({len(self.printers)}):")
        lines.extend(f"  - {p}" for p in self.printers)
        lines.append("")
        lines.append(f"数据目录: {self.data_dir}")
        lines.append(f"日志目录: {self.logs_dir}")
        problems = self.problems()
        if problems:
            lines.append("")
            lines.append("需要注意:")
            lines.extend(f"  ! {p}" for p in problems)
        return "\n".join(lines)


def collect_environment(preference: str = "auto", deep: bool = False) -> EnvironmentReport:
    """收集环境信息；deep=True 时真正启动每个 COM 组件（较慢，约数秒到数十秒）"""
    from ..core.printer_environment import get_default_printer
    report = EnvironmentReport(deep=deep)
    report.windows = windows_version()
    report.python = f"{platform.python_version()} ({platform.architecture()[0]})"
    backends = default_backends()
    statuses: Dict[str, Dict[OfficeKind, ComponentStatus]] = {}
    if deep:
        with office_automation(1, "正在打印或统计页数，请完成后再进行深度检测"), com_initialized():
            for backend in backends:
                statuses[backend.key] = {}
                for kind in OfficeKind:
                    quick = backend.check_component(kind)
                    statuses[backend.key][kind] = (backend.deep_check_component(kind)
                                                   if quick.registered else quick)
    else:
        for backend in backends:
            statuses[backend.key] = backend.check_all()
    report.statuses = statuses
    report.selection = select_backends(backends, preference, statuses)
    sumatra = path_utils.get_sumatra_pdf_path()
    report.sumatra_path = str(sumatra)
    report.sumatra_ok = sumatra.is_file()
    report.printers = list_printers()
    report.default_printer = get_default_printer() or ""
    names = set()
    for backend in backends:
        names |= backend.all_process_names
    report.running_office = process_utils.list_running(names)
    report.data_dir = str(path_utils.get_data_root())
    report.logs_dir = str(path_utils.get_logs_dir())
    return report
