"""
真实环境端到端打印验证（Windows + WPS Office）

使用真实的 PrintController、真实 WPS COM、真实 SumatraPDF，把 tests/samples 中的文件
打印到一台“文件端口虚拟 PDF 打印机”（见 scripts/setup_test_pdf_printer.ps1），
然后检查每个任务的输出 PDF（页数、页面方向）、WPS 进程是否退出、默认打印机是否恢复。

安全：默认拒绝向名称中不含 PDF/XPS 的打印机发送作业，避免误用真实物理打印机消耗纸张。

用法（管理员 PowerShell 中先创建测试打印机）:
    powershell -ExecutionPolicy Bypass -File scripts\\setup_test_pdf_printer.ps1 -OutputFile C:\\bdp_test\\out.pdf
    python scripts\\e2e_real_print.py --printer "BDP Test PDF" --port-file C:\\bdp_test\\out.pdf --native
"""
import argparse
import json
import shutil
import sys
import threading
import time
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
from src.core.models import Document, PrintSettings, PrintStatus  # noqa: E402
from src.core.print_controller import PrintController  # noqa: E402
from src.core.printer_environment import get_default_printer  # noqa: E402
from src.utils.logging_setup import setup_logging  # noqa: E402

WPS_PROCS = {"wps.exe", "et.exe", "wpp.exe"}
DEFAULT_FILES = ["sample.docx", "sample.xlsx", "sample.pptx", "sample.pdf", "sample.png", "sample.jpg",
                 "sample.bmp", "sample.tiff", "sample.txt"]


def pdf_info(path: Path) -> dict:
    from PyPDF2 import PdfReader
    reader = PdfReader(str(path))
    pages = []
    for page in reader.pages:
        box = page.mediabox
        w, h = float(box.width), float(box.height)
        rotate = int(page.get("/Rotate", 0) or 0) % 180
        if rotate:
            w, h = h, w
        pages.append({"w_mm": round(w / 72 * 25.4), "h_mm": round(h / 72 * 25.4),
                      "orientation": "landscape" if w > h else "portrait"})
    return {"pages": len(pages), "page_sizes": pages}


def make_native_files(samples: Path, out_dir: Path, log) -> list:
    """用 WPS 把 docx/xlsx/pptx 另存为 WPS 原生格式 .wps/.et/.dps"""
    from src.backends.office_backend import OfficeKind, OfficeSession, com_initialized
    from src.backends.wps_backend import WpsBackend
    from src.backends import office_ops as ops
    created = []
    backend = WpsBackend()
    jobs = [(OfficeKind.WRITER, "sample.docx", "原生 文字.wps"),
            (OfficeKind.SPREADSHEET, "sample.xlsx", "原生 表格.et"),
            (OfficeKind.PRESENTATION, "sample.pptx", "原生 演示.dps")]
    with com_initialized():
        for kind, src_name, dest_name in jobs:
            dest = out_dir / dest_name
            session = OfficeSession(backend, kind, log=log)
            try:
                session.start()
                app = session.app
                if kind == OfficeKind.WRITER:
                    doc = ops.open_writer_document(app, samples / src_name, log)
                    doc.SaveAs2(str(dest))
                    ops.close_writer_document(doc, log)
                elif kind == OfficeKind.SPREADSHEET:
                    wb = ops.open_workbook(app, samples / src_name, log)
                    wb.SaveAs(str(dest))
                    ops.close_workbook(wb, log)
                else:
                    pres = ops.open_presentation(app, samples / src_name, log)
                    pres.SaveAs(str(dest))
                    ops.close_presentation(pres, log)
                if dest.exists():
                    header = dest.read_bytes()[:8].hex()
                    log(f"已生成 {dest.name}（{dest.stat().st_size} 字节，文件头 {header}）")
                    created.append(dest)
                else:
                    log(f"!! WPS 未生成 {dest.name}")
            except Exception as e:  # noqa: BLE001 - 报告后继续
                log(f"!! 生成 {dest_name} 失败: {e}")
            finally:
                session.close()
    return created


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--printer", default="BDP Test PDF")
    parser.add_argument("--port-file", type=Path, required=True, help="测试打印机端口对应的输出 PDF 文件路径")
    parser.add_argument("--samples", type=Path, default=ROOT / "tests" / "samples")
    parser.add_argument("--out-dir", type=Path, default=ROOT / "e2e_output")
    parser.add_argument("--native", action="store_true", help="用 WPS 生成 .wps/.et/.dps 并一起打印")
    parser.add_argument("--copies", type=int, default=1)
    parser.add_argument("--allow-physical-printer", action="store_true")
    parser.add_argument("--wait", type=float, default=120, help="每个作业等待输出文件的秒数")
    args = parser.parse_args()

    if sys.platform != "win32":
        print("此脚本只能在 Windows 上运行")
        return 2
    if not any(k in args.printer.upper() for k in ("PDF", "XPS")) and not args.allow_physical_printer:
        print(f"拒绝向可能是物理打印机的 {args.printer!r} 发送作业（如确认请加 --allow-physical-printer）")
        return 2

    setup_logging(redirect_std=False)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    work = args.out_dir / "输入 文件（中文 路径）"
    work.mkdir(exist_ok=True)

    def log(msg):
        print(msg, flush=True)

    files = []
    for name in DEFAULT_FILES:
        dest = work / name.replace("sample", "样例 文件")
        shutil.copy(args.samples / name, dest)
        files.append(dest)
    corrupt = work / "损坏的文档.docx"
    corrupt.write_bytes(b"this is not a docx file")
    files.append(corrupt)
    if args.native:
        files.extend(make_native_files(args.samples, work, log))

    initial_default = get_default_printer()
    before = process_utils.snapshot(WPS_PROCS)
    log(f"初始默认打印机: {initial_default}；已有 WPS 进程: {sorted(before)}")

    results = {}
    job_start = {}
    lock = threading.Lock()

    def on_document(doc: Document):
        port = args.port_file
        if doc.print_status == PrintStatus.PRINTING:
            with lock:
                job_start[doc.id] = time.time()
            if port.exists():
                port.unlink()
            return
        if doc.print_status not in (PrintStatus.COMPLETED, PrintStatus.ERROR):
            return
        entry = {"file": doc.file_name, "status": doc.print_status.value, "error": doc.last_error,
                 "backend": doc.last_backend}
        if doc.print_status == PrintStatus.COMPLETED:
            deadline = time.time() + args.wait
            last_size = -1
            stable_since = None
            while time.time() < deadline:
                if port.exists():
                    size = port.stat().st_size
                    if size > 0 and size == last_size:
                        stable_since = stable_since or time.time()
                        if time.time() - stable_since > 2:
                            break
                    else:
                        stable_since = None
                    last_size = size
                time.sleep(0.5)
            if port.exists() and port.stat().st_size > 0:
                target = args.out_dir / f"{len(results) + 1:02d}_{doc.file_name}.pdf"
                for _ in range(20):
                    try:
                        shutil.move(str(port), str(target))
                        break
                    except OSError:
                        time.sleep(0.5)
                entry["output"] = str(target)
                try:
                    entry.update(pdf_info(target))
                except Exception as e:  # noqa: BLE001
                    entry["pdf_error"] = str(e)
            else:
                entry["output"] = None
        with lock:
            results[doc.file_name] = entry
        log(f"[{entry['status']}] {doc.file_name} backend={entry['backend']} "
            f"pages={entry.get('pages')} error={entry['error']}")

    controller = PrintController(inter_job_delay=1.0)
    controller.set_document_callback(on_document)
    controller.set_print_settings(PrintSettings(printer_name=args.printer, copies=args.copies))
    controller.add_documents_to_queue([Document(file_path=f) for f in files])
    summary = controller.start_batch_print().result(timeout=3600)
    controller.shutdown()

    time.sleep(3)
    after = process_utils.snapshot(WPS_PROCS)
    leftovers = sorted(set(after) - set(before))
    final_default = get_default_printer()
    report = {
        "printer": args.printer,
        "summary": {"success": summary.success, "failed": summary.failed, "text": summary.text()},
        "warnings": summary.warnings,
        "default_printer_before": initial_default,
        "default_printer_after": final_default,
        "default_printer_restored": initial_default == final_default,
        "leftover_wps_processes": leftovers,
        "results": results,
    }
    (args.out_dir / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n================ 结果 ================")
    for entry in results.values():
        print(f"{entry['status']:9} pages={str(entry.get('pages')):4} {entry['file']:28} {entry['backend']}"
              + (f"  错误: {entry['error']}" if entry['error'] else ""))
    print(f"批次: {summary.text()}")
    print(f"默认打印机恢复: {report['default_printer_restored']} ({initial_default} -> {final_default})")
    print(f"残留 WPS 进程: {leftovers or '无'}")

    ok = report["default_printer_restored"] and not leftovers
    for entry in results.values():
        expected_fail = entry["file"] == corrupt.name
        if expected_fail:
            ok = ok and entry["status"] == "error"
        else:
            ok = ok and entry["status"] == "completed" and bool(entry.get("pages"))
    print("E2E RESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
