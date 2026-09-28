"""
WPS COM 冒烟测试（不打印任何内容）

对 KWPS.Application / KET.Application / KWPP.Application 依次执行：
    注册表检查 → 创建 COM 对象 → Visible=False → 读取版本 → Quit → 确认进程退出

用法:
    python scripts/test_wps_com.py                 # 仅 WPS
    python scripts/test_wps_com.py --msoffice      # 同时测试 Microsoft Office（可选）
    python scripts/test_wps_com.py --timeout 120

输出示例:
    WPS Writer COM: OK
    WPS Spreadsheet COM: OK
    WPS Presentation COM: OK

退出码: 0 = 全部通过, 1 = 至少一个失败, 2 = 非 Windows 环境
"""
import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

# 非中文系统的控制台代码页无法显示中文时，用 ? 代替而不是崩溃
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

from src.backends import process_utils  # noqa: E402
from src.backends.office_backend import OfficeKind, com_initialized  # noqa: E402
from src.backends.ms_office_backend import MsOfficeBackend  # noqa: E402
from src.backends.wps_backend import WpsBackend  # noqa: E402

LABELS = {
    ("wps", OfficeKind.WRITER): "WPS Writer COM",
    ("wps", OfficeKind.SPREADSHEET): "WPS Spreadsheet COM",
    ("wps", OfficeKind.PRESENTATION): "WPS Presentation COM",
    ("msoffice", OfficeKind.WRITER): "Microsoft Word COM",
    ("msoffice", OfficeKind.SPREADSHEET): "Microsoft Excel COM",
    ("msoffice", OfficeKind.PRESENTATION): "Microsoft PowerPoint COM",
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--timeout", type=float, default=90, help="每个组件的超时秒数")
    parser.add_argument("--msoffice", action="store_true", help="同时测试 Microsoft Office")
    parser.add_argument("--json", type=Path, help="把结果写入 JSON 文件")
    args = parser.parse_args()

    if sys.platform != "win32":
        print("此脚本只能在 Windows 上运行")
        return 2

    backends = [WpsBackend()] + ([MsOfficeBackend()] if args.msoffice else [])
    results = []
    all_ok = True
    with com_initialized():
        for backend in backends:
            before = process_utils.snapshot(backend.all_process_names)
            for kind in OfficeKind:
                status = backend.deep_check_component(kind, timeout=args.timeout)
                label = LABELS[(backend.key, kind)]
                ok = bool(status.deep_ok)
                if backend.key == "wps":
                    all_ok = all_ok and ok
                print(f"{label}: {'OK' if ok else 'FAIL'}"
                      + (f" (version {status.version})" if status.version else "")
                      + ("" if ok else f" - ProgID {status.progid}: {status.error}"))
                for line in status.detail_lines()[1:]:
                    print(f"    {line.strip()}")
                results.append({"component": label, "progid": status.progid, "ok": ok,
                                 "version": status.version, "clsid": status.clsid,
                                 "server": status.server_path, "error": status.error,
                                 "detail": status.deep_detail})
            after = process_utils.snapshot(backend.all_process_names)
            leftovers = sorted(set(after) - set(before))
            if leftovers:
                print(f"!! {backend.display_name} 测试后残留进程: {leftovers}")
                if backend.key == "wps":
                    all_ok = False
            else:
                print(f"{backend.display_name}: 测试后无残留进程")

    if args.json:
        args.json.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if all_ok else 1


if __name__ == "__main__":
    code = main()
    if getattr(sys, "frozen", False) and len(sys.argv) == 1:
        # 双击运行打包后的诊断工具时，保留窗口以便查看结果
        try:
            input("\n按回车键退出...")
        except EOFError:
            pass
    sys.exit(code)
