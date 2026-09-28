"""
Office 文档操作（打开 / 打印 / 页数统计 / 关闭）

同一套实现同时服务 WPS（KWPS/KET/KWPP）与 Microsoft Office，
WPS 与 MS Office 的差异通过 call_with_fallbacks 的多级降级调用处理：
    命名参数 → 位置参数 → 最简调用 → 循环打印份数
"""
import time
from pathlib import Path
from typing import Callable, List, Optional

from src.core.models import Orientation, PrintSettings
from .com_utils import (call_with_fallbacks, com_retry, describe_com_error, is_signature_error,
                        looks_like_password_error, safe_get, safe_set)

LogFunc = Callable[[str], None]

# 打开文件时传入的"假密码"：对未加密文件无影响；对加密文件会直接报错而不是弹出密码框卡住
DUMMY_PASSWORD = "bdp-no-password-7f3a"

MSO_TRUE = -1
MSO_FALSE = 0

# 纸张代码：Word 使用 WdPaperSize，Excel/WPS 表格使用与 DEVMODE 相同的 XlPaperSize
WORD_PAPER_CODES = {"Letter": 2, "Legal": 4, "A3": 6, "A4": 7, "A5": 9, "B5": 11}
EXCEL_PAPER_CODES = {"Letter": 1, "Legal": 5, "A3": 8, "A4": 9, "A5": 11, "B5": 13}


class OfficeDocumentError(RuntimeError):
    """文档级错误（带有面向用户的中文说明）"""


def _noop(_msg: str) -> None:
    pass


def printer_matches(active_printer: Optional[str], target: str) -> bool:
    """ActivePrinter 可能是 "打印机名" 或 "打印机名 on Ne01:" / "打印机名 在 Ne01: 上" 等形式"""
    if not active_printer or not target:
        return False
    active = str(active_printer).strip().lower()
    target = target.strip().lower()
    return active == target or active.startswith(target + " ")


def try_set_active_printer(obj, target: str, log: LogFunc, label: str) -> bool:
    """
    尝试设置 ActivePrinter 并读回验证。
    返回 True 表示已确认目标打印机生效；False 表示 API 不可写或读回不一致
    （批量打印时系统默认打印机已临时切换为目标打印机，因此仍会打印到正确的打印机）。
    """
    if not target:
        return False
    current = safe_get(obj, "ActivePrinter")
    if printer_matches(current, target):
        log(f"{label}.ActivePrinter 已是目标打印机: {current}")
        return True
    try:
        setattr(obj, "ActivePrinter", target)
    except Exception as e:  # noqa: BLE001
        log(f"{label}.ActivePrinter 不可写（{describe_com_error(e)}），依赖系统默认打印机切换")
        return False
    after = safe_get(obj, "ActivePrinter")
    ok = printer_matches(after, target)
    log(f"{label}.ActivePrinter 设置{'成功' if ok else '后读回不一致'}: {after}")
    return ok


def _open_with_password_guard(attempts_with_pwd, attempts_plain, log: LogFunc, what: str):
    """
    先用"假密码"方式打开；若报密码错误直接判定为加密文件；
    若是其它非签名错误，再尝试不带密码的方式打开一次。
    """
    try:
        doc, method = call_with_fallbacks(attempts_with_pwd, log)
        return doc, method
    except Exception as first:  # noqa: BLE001
        if looks_like_password_error(first):
            raise OfficeDocumentError(f"文件受密码保护，无法自动打印: {describe_com_error(first)}") from first
        log(f"{what} 打开失败（{describe_com_error(first)}），尝试不带密码参数重新打开")
        try:
            return call_with_fallbacks(attempts_plain, log)
        except Exception as second:  # noqa: BLE001
            if looks_like_password_error(second):
                raise OfficeDocumentError(f"文件受密码保护，无法自动打印: {describe_com_error(second)}") from second
            raise OfficeDocumentError(f"无法打开文件: {describe_com_error(second)}") from second


def _wait_background_printing(app, log: LogFunc, timeout: float = 300.0) -> None:
    """等待后台打印完成（Word/WPS 文字 BackgroundPrintingStatus）"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        status = safe_get(app, "BackgroundPrintingStatus", 0)
        try:
            if int(status or 0) <= 0:
                return
        except (TypeError, ValueError):
            return
        time.sleep(0.5)
    log("等待后台打印完成超时，继续执行")


# ======================================================================
# 文字（Word / WPS 文字）
# ======================================================================

def open_writer_document(app, path: Path, log: LogFunc = _noop):
    p = str(path)
    docs = app.Documents
    with_pwd = [
        ("Documents.Open 命名参数", lambda: docs.Open(FileName=p, ConfirmConversions=False, ReadOnly=True,
                                                   AddToRecentFiles=False, PasswordDocument=DUMMY_PASSWORD,
                                                   NoEncodingDialog=True)),
        ("Documents.Open 位置参数", lambda: docs.Open(p, False, True, False, DUMMY_PASSWORD)),
    ]
    plain = [
        ("Documents.Open 只读", lambda: docs.Open(p, False, True, False)),
        ("Documents.Open 仅路径", lambda: docs.Open(p)),
    ]
    doc, method = _open_with_password_guard(with_pwd, plain, log, "文字文档")
    if doc is None:
        raise OfficeDocumentError("打开文档失败：返回空对象（文件可能已加密或损坏）")
    log(f"已打开文字文档（{method}）")
    return doc


def close_writer_document(doc, log: LogFunc = _noop) -> None:
    if doc is None:
        return
    try:
        call_with_fallbacks([("Close(SaveChanges=0)", lambda: doc.Close(SaveChanges=0)),
                             ("Close(0)", lambda: doc.Close(0)),
                             ("Close()", lambda: doc.Close())], log)
    except Exception as e:  # noqa: BLE001 - 关闭失败由会话清理兜底
        log(f"关闭文字文档失败: {describe_com_error(e)}")


def apply_writer_page_setup(doc, settings: PrintSettings, log: LogFunc) -> List[str]:
    """仅在用户选择"强制套用纸张/方向"时修改文档页面设置（不保存）"""
    warnings = []
    page_setup = safe_get(doc, "PageSetup")
    if page_setup is None:
        return ["无法读取文档页面设置，已使用文档自身设置"]
    if settings.orientation != Orientation.AUTO:
        value = 1 if settings.orientation == Orientation.LANDSCAPE else 0
        if not safe_set(page_setup, "Orientation", value, log):
            warnings.append("无法修改文档方向，已使用文档自身设置")
    code = WORD_PAPER_CODES.get(settings.paper_size)
    if code is None:
        warnings.append(f"纸张 {settings.paper_size} 无法映射到文字文档纸张代码，已使用文档自身设置")
    elif not safe_set(page_setup, "PaperSize", code, log):
        warnings.append(f"无法将文档纸张设置为 {settings.paper_size}，已使用文档自身设置")
    return warnings


def writer_print(app, doc, copies: int, log: LogFunc = _noop) -> str:
    """打印文字文档，返回实际使用的调用方式"""
    copies = max(1, int(copies))
    attempts = [
        ("PrintOut(Background=False, Copies, Collate)",
         lambda: doc.PrintOut(Background=False, Copies=copies, Collate=True)),
        ("PrintOut 位置参数", lambda: doc.PrintOut(False, False, 0, "", "", "", 0, copies)),
    ]
    try:
        _, method = call_with_fallbacks(attempts, log)
    except Exception as e:  # noqa: BLE001
        if not is_signature_error(e):
            raise
        log("PrintOut 不支持份数参数，改为逐份调用 PrintOut()")
        for _ in range(copies):
            com_retry(lambda: doc.PrintOut())
            _wait_background_printing(app, log)
        method = f"PrintOut() x{copies}"
    _wait_background_printing(app, log)
    return method


def writer_count_pages(doc, log: LogFunc = _noop) -> int:
    attempts = [
        ("ComputeStatistics(wdStatisticPages)", lambda: doc.ComputeStatistics(2)),
        ("Content.Information(wdNumberOfPagesInDocument)", lambda: doc.Content.Information(4)),
    ]
    pages, method = call_with_fallbacks(attempts, log)
    pages = int(pages)
    if pages < 0:
        raise OfficeDocumentError("获取到无效的页数")
    log(f"文字文档页数 {pages}（{method}）")
    return pages


# ======================================================================
# 表格（Excel / WPS 表格）
# ======================================================================

XL_SHEET_VISIBLE = -1


def open_workbook(app, path: Path, log: LogFunc = _noop):
    p = str(path)
    books = app.Workbooks
    with_pwd = [
        ("Workbooks.Open 命名参数", lambda: books.Open(Filename=p, UpdateLinks=0, ReadOnly=True,
                                                   Password=DUMMY_PASSWORD, IgnoreReadOnlyRecommended=True,
                                                   AddToMru=False)),
    ]
    plain = [
        ("Workbooks.Open 位置参数", lambda: books.Open(p, 0, True)),
        ("Workbooks.Open 仅路径", lambda: books.Open(p)),
    ]
    wb, method = _open_with_password_guard(with_pwd, plain, log, "工作簿")
    if wb is None:
        raise OfficeDocumentError("打开工作簿失败：返回空对象")
    log(f"已打开工作簿（{method}）")
    return wb


def close_workbook(wb, log: LogFunc = _noop) -> None:
    if wb is None:
        return
    try:
        call_with_fallbacks([("Close(SaveChanges=False)", lambda: wb.Close(SaveChanges=False)),
                             ("Close(False)", lambda: wb.Close(False)),
                             ("Close()", lambda: wb.Close())], log)
    except Exception as e:  # noqa: BLE001
        log(f"关闭工作簿失败: {describe_com_error(e)}")


def iter_worksheets(wb):
    sheets = wb.Worksheets
    count = int(safe_get(sheets, "Count", 0) or 0)
    for i in range(1, count + 1):
        yield sheets.Item(i)


def worksheet_is_visible(ws) -> bool:
    return safe_get(ws, "Visible", XL_SHEET_VISIBLE) == XL_SHEET_VISIBLE


def worksheet_has_content(ws) -> bool:
    """粗略判断工作表是否有可打印内容（单元格内容、图形或打印区域）"""
    if str(safe_get(safe_get(ws, "PageSetup"), "PrintArea", "") or ""):
        return True
    shapes = safe_get(ws, "Shapes")
    if shapes is not None and int(safe_get(shapes, "Count", 0) or 0) > 0:
        return True
    used = safe_get(ws, "UsedRange")
    if used is None:
        return True  # 无法判断时按有内容处理
    address = str(safe_get(used, "Address", "") or "")
    if address and address.replace("$", "") != "A1":
        return True
    try:
        value = ws.Range("A1").Value
    except Exception:  # noqa: BLE001 - 无法判断时按有内容处理
        return True
    return value not in (None, "")


def worksheet_page_count(ws, log: LogFunc = _noop) -> int:
    """单个工作表的打印页数：优先 PageSetup.Pages.Count，其次分页符推算"""
    try:
        pages = int(ws.PageSetup.Pages.Count)
        if pages >= 0:
            return pages
    except Exception:  # noqa: BLE001 - WPS 旧版本可能不支持 Pages
        pass
    try:
        h_breaks = int(ws.HPageBreaks.Count)
        v_breaks = int(ws.VPageBreaks.Count)
        return (h_breaks + 1) * (v_breaks + 1)
    except Exception as e:  # noqa: BLE001
        log(f"工作表 {safe_get(ws, 'Name', '?')} 页数统计失败，按 1 页计算: {describe_com_error(e)}")
        return 1


def spreadsheet_count_pages(wb, log: LogFunc = _noop) -> int:
    total = 0
    details = []
    for ws in iter_worksheets(wb):
        name = safe_get(ws, "Name", "?")
        if not worksheet_is_visible(ws):
            details.append(f"{name}:隐藏")
            continue
        if not worksheet_has_content(ws):
            details.append(f"{name}:空")
            continue
        pages = worksheet_page_count(ws, log)
        details.append(f"{name}:{pages}")
        total += pages
    log(f"工作簿页数 {total}（{', '.join(details)}）")
    return total


def apply_spreadsheet_page_setup(wb, settings: PrintSettings, log: LogFunc) -> List[str]:
    """强制套用纸张/方向到每个可见工作表（不保存）"""
    warnings = []
    code = EXCEL_PAPER_CODES.get(settings.paper_size)
    if code is None:
        warnings.append(f"纸张 {settings.paper_size} 无法映射到表格纸张代码，已使用工作表自身设置")
    for ws in iter_worksheets(wb):
        if not worksheet_is_visible(ws):
            continue
        name = safe_get(ws, "Name", "?")
        page_setup = safe_get(ws, "PageSetup")
        if page_setup is None:
            warnings.append(f"工作表 {name} 无法读取页面设置")
            continue
        if settings.orientation != Orientation.AUTO:
            value = 2 if settings.orientation == Orientation.LANDSCAPE else 1
            if not safe_set(page_setup, "Orientation", value, log):
                warnings.append(f"工作表 {name} 方向设置失败，已使用自身设置")
        if code is not None and not safe_set(page_setup, "PaperSize", code, log):
            warnings.append(f"工作表 {name} 纸张 {settings.paper_size} 设置失败（打印机驱动可能不支持），已使用自身设置")
    return warnings


def spreadsheet_print(app, wb, copies: int, log: LogFunc = _noop) -> str:
    """打印整个工作簿（所有可见工作表，遵循各表打印区域）"""
    copies = max(1, int(copies))
    printable = [ws for ws in iter_worksheets(wb) if worksheet_is_visible(ws) and worksheet_has_content(ws)]
    chart_sheets = int(safe_get(safe_get(wb, "Charts"), "Count", 0) or 0)
    if not printable and chart_sheets == 0:
        raise OfficeDocumentError("工作簿中没有可打印的内容（所有工作表为空或隐藏）")

    attempts = [
        ("Workbook.PrintOut(Copies, Collate)", lambda: wb.PrintOut(Copies=copies, Collate=True)),
        ("Workbook.PrintOut 位置参数", lambda: wb.PrintOut(1, 32767, copies)),
    ]
    try:
        _, method = call_with_fallbacks(attempts, log)
        return method
    except Exception as e:  # noqa: BLE001
        if not is_signature_error(e):
            raise
    # 最后降级：逐份、逐表打印
    log("Workbook.PrintOut 不支持份数参数，改为逐份逐表调用 Worksheet.PrintOut()")
    for _ in range(copies):
        for ws in printable:
            com_retry(lambda ws=ws: ws.PrintOut())
    return f"Worksheet.PrintOut() x{copies} x{len(printable)}表"


# ======================================================================
# 演示（PowerPoint / WPS 演示）
# ======================================================================

PP_PRINT_COLOR = 1
PP_PRINT_BLACK_AND_WHITE = 2
PP_PRINT_OUTPUT_SLIDES = 1
PP_PRINT_ALL = 1


def open_presentation(app, path: Path, log: LogFunc = _noop):
    p = str(path)
    presentations = app.Presentations
    attempts = [
        ("Presentations.Open 命名参数(无窗口)",
         lambda: presentations.Open(FileName=p, ReadOnly=MSO_TRUE, Untitled=MSO_FALSE, WithWindow=MSO_FALSE)),
        ("Presentations.Open 位置参数(无窗口)", lambda: presentations.Open(p, MSO_TRUE, MSO_FALSE, MSO_FALSE)),
        ("Presentations.Open 仅路径", lambda: presentations.Open(p)),
    ]
    try:
        pres, method = call_with_fallbacks(attempts, log)
    except Exception as e:  # noqa: BLE001
        if looks_like_password_error(e):
            raise OfficeDocumentError(f"文件受密码保护，无法自动打印: {describe_com_error(e)}") from e
        raise OfficeDocumentError(f"无法打开演示文稿: {describe_com_error(e)}") from e
    if pres is None:
        raise OfficeDocumentError("打开演示文稿失败：返回空对象")
    log(f"已打开演示文稿（{method}）")
    return pres


def close_presentation(pres, log: LogFunc = _noop) -> None:
    if pres is None:
        return
    try:
        safe_set(pres, "Saved", MSO_TRUE)
        com_retry(lambda: pres.Close())
    except Exception as e:  # noqa: BLE001
        log(f"关闭演示文稿失败: {describe_com_error(e)}")


def presentation_count_slides(pres, log: LogFunc = _noop) -> int:
    count = int(pres.Slides.Count)
    log(f"演示文稿幻灯片数 {count}")
    return count


def presentation_print(pres, settings: PrintSettings, log: LogFunc = _noop) -> dict:
    """
    打印演示文稿。

    Returns:
        {"method": 调用方式, "active_printer_writable": bool}
    """
    copies = max(1, int(settings.copies))
    options = safe_get(pres, "PrintOptions")
    active_printer_ok = False
    copies_via_options = False
    if options is not None:
        safe_set(options, "PrintInBackground", MSO_FALSE, log)
        active_printer_ok = try_set_active_printer(options, settings.printer_name, log, "PrintOptions")
        safe_set(options, "OutputType", PP_PRINT_OUTPUT_SLIDES)
        safe_set(options, "RangeType", PP_PRINT_ALL)
        safe_set(options, "FitToPage", MSO_TRUE)
        safe_set(options, "PrintColorType", PP_PRINT_COLOR if settings.color else PP_PRINT_BLACK_AND_WHITE, log)
        safe_set(options, "Collate", MSO_TRUE)
        if safe_set(options, "NumberOfCopies", copies, log):
            copies_via_options = int(safe_get(options, "NumberOfCopies", 0) or 0) == copies

    if copies_via_options:
        com_retry(lambda: pres.PrintOut())
        method = f"PrintOptions.NumberOfCopies={copies} + PrintOut()"
    else:
        attempts = [("PrintOut(Copies, Collate)", lambda: pres.PrintOut(Copies=copies, Collate=MSO_TRUE)),
                    ("PrintOut 位置参数", lambda: pres.PrintOut(-1, -1, "", copies, MSO_TRUE))]
        try:
            _, method = call_with_fallbacks(attempts, log)
        except Exception as e:  # noqa: BLE001
            if not is_signature_error(e):
                raise
            log("PrintOut 不支持份数参数，改为逐份调用 PrintOut()")
            for _ in range(copies):
                com_retry(lambda: pres.PrintOut())
            method = f"PrintOut() x{copies}"
    return {"method": method, "active_printer_writable": active_printer_ok}
