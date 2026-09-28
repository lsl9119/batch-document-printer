"""
直接通过 Windows GDI 打印（不依赖任何外部程序）

- 文本：自动换行、分页，中文字体（宋体/微软雅黑）
- 图片：SumatraPDF 不可用时的备用方案（Pillow ImageWin）

打印参数（纸张/方向/双面/色彩）来自打印机的每用户默认 DEVMODE，
批量打印时由 PrinterEnvironment 预先写入。
"""
import logging
from pathlib import Path
from typing import Callable, List

from ..core.models import Orientation, PrintSettings, ScalingMode
from ..core.print_context import PrintJobError

logger = logging.getLogger("bdp.gdi")

TAB_SIZE = 4
# 依次尝试的字体（GDI 找不到字体时会静默替换，因此选中后需核对实际字体名）
FONT_CANDIDATES = ("SimSun", "NSimSun", "Microsoft YaHei", "SimHei", "DengXian", "Noto Sans CJK SC",
                   "Source Han Sans SC", "Arial Unicode MS")
FALLBACK_FONT = "Courier New"
FONT_POINTS = 10.5
MARGIN_MM = 15


def wrap_line(line: str, measure: Callable[[str], int], max_width: int) -> List[str]:
    """按像素宽度折行（二分查找最长可容纳前缀；西文优先在空格处断行）"""
    if max_width <= 0:
        return [line]
    out: List[str] = []
    while line:
        if measure(line) <= max_width:
            out.append(line)
            break
        lo, hi = 1, len(line)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if measure(line[:mid]) <= max_width:
                lo = mid
            else:
                hi = mid - 1
        cut = max(1, lo)
        space = line.rfind(" ", 0, cut)
        if space > cut * 0.6:
            cut = space + 1
        out.append(line[:cut].rstrip())
        line = line[cut:]
    return out or [""]


def layout_text(text: str, measure: Callable[[str], int], max_width: int,
                lines_per_page: int) -> List[List[str]]:
    """把文本排成若干页（换页符 \\f 强制分页）"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    pages: List[List[str]] = []
    lines_per_page = max(1, lines_per_page)
    for chunk in text.split("\f"):
        page: List[str] = []
        for raw in chunk.split("\n"):
            for line in wrap_line(raw.expandtabs(TAB_SIZE), measure, max_width):
                if len(page) >= lines_per_page:
                    pages.append(page)
                    page = []
                page.append(line)
        pages.append(page)
    # 去掉末尾的空白页
    while len(pages) > 1 and not any(l.strip() for l in pages[-1]):
        pages.pop()
    return pages or [[]]


def _printer_dc(printer_name: str):
    import win32ui  # type: ignore
    dc = win32ui.CreateDC()
    try:
        dc.CreatePrinterDC(printer_name)
    except Exception as e:  # noqa: BLE001
        raise PrintJobError(f"无法连接打印机 {printer_name}: {e}") from e
    return dc


def _margins(dc):
    import win32con  # type: ignore
    dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
    dpi_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)
    off_x = dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
    off_y = dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)
    mx = max(0, int(MARGIN_MM / 25.4 * dpi_x) - off_x)
    my = max(0, int(MARGIN_MM / 25.4 * dpi_y) - off_y)
    return dpi_x, dpi_y, mx, my


def _select_font(dc, win32ui, height: int):
    """
    选择第一个系统中真实存在的中文字体。

    Returns:
        (字体名, 字体对象)。字体对象必须在打印结束前保持引用，否则会被提前释放。
    """
    last = (FALLBACK_FONT, None)
    for name in FONT_CANDIDATES + (FALLBACK_FONT,):
        try:
            font = win32ui.CreateFont({"name": name, "height": height, "weight": 400, "charset": 1})
            dc.SelectObject(font)
        except Exception:  # noqa: BLE001 - 尝试下一个字体
            continue
        last = (name, font)
        try:
            actual = dc.GetTextFace()
        except Exception:  # noqa: BLE001 - 无法核对时直接使用
            return name, font
        if actual and actual.lower() == name.lower():
            return name, font
    return last


def print_text(printer_name: str, title: str, text: str, copies: int, log=None) -> int:
    """GDI 打印文本，返回每份的页数"""
    import win32con  # type: ignore
    import win32ui  # type: ignore
    log = log or (lambda m: logger.info(m))
    dc = _printer_dc(printer_name)
    started = False
    try:
        width = dc.GetDeviceCaps(win32con.HORZRES)
        height = dc.GetDeviceCaps(win32con.VERTRES)
        dpi_x, dpi_y, mx, my = _margins(dc)
        font_name, font = _select_font(dc, win32ui, -int(FONT_POINTS * dpi_y / 72))
        log(f"GDI 文本打印字体: {font_name}")
        line_height = max(1, int(dc.GetTextExtent("国Ag")[1] * 1.3))
        lines_per_page = max(1, (height - 2 * my) // line_height)
        pages = layout_text(text, lambda s: dc.GetTextExtent(s)[0], width - 2 * mx, lines_per_page)
        dc.StartDoc(title)
        started = True
        for _ in range(max(1, copies)):
            for page in pages:
                dc.StartPage()
                for i, line in enumerate(page):
                    if line:
                        dc.TextOut(mx, my + i * line_height, line)
                dc.EndPage()
        dc.EndDoc()
        started = False
        log(f"GDI 文本打印完成：每份 {len(pages)} 页 x {copies} 份")
        return len(pages)
    except PrintJobError:
        raise
    except Exception as e:  # noqa: BLE001
        raise PrintJobError(f"GDI 文本打印失败: {e}") from e
    finally:
        if started:
            try:
                dc.AbortDoc()
            except Exception:  # noqa: BLE001
                pass
        dc.DeleteDC()


def fit_rect(img_w: int, img_h: int, area_w: int, area_h: int, scaling: ScalingMode,
             img_dpi: float, dev_dpi: float):
    """计算图片在可打印区域中的目标尺寸 (w, h)"""
    natural_w = img_w / img_dpi * dev_dpi
    natural_h = img_h / img_dpi * dev_dpi
    fit_scale = min(area_w / natural_w, area_h / natural_h)
    if scaling == ScalingMode.FIT:
        scale = fit_scale
    elif scaling == ScalingMode.SHRINK:
        scale = min(1.0, fit_scale)
    else:
        scale = 1.0
    return max(1, int(natural_w * scale)), max(1, int(natural_h * scale))


def print_image(printer_name: str, file_path: Path, settings: PrintSettings, log=None) -> int:
    """GDI 打印图片（多页 TIFF 逐页），返回每份页数"""
    import win32con  # type: ignore
    from PIL import Image, ImageSequence, ImageWin  # type: ignore
    log = log or (lambda m: logger.info(m))
    dc = _printer_dc(printer_name)
    started = False
    try:
        width = dc.GetDeviceCaps(win32con.HORZRES)
        height = dc.GetDeviceCaps(win32con.VERTRES)
        dpi_x, dpi_y, mx, my = _margins(dc)
        area_w, area_h = width - 2 * mx, height - 2 * my
        with Image.open(file_path) as img:
            frames = [f.copy() for f in ImageSequence.Iterator(img)]
            img_dpi = float((img.info.get("dpi") or (96, 96))[0] or 96)
        dc.StartDoc(file_path.name)
        started = True
        for _ in range(max(1, settings.copies)):
            for frame in frames:
                frame = frame.convert("RGB") if settings.color else frame.convert("L").convert("RGB")
                if settings.orientation == Orientation.AUTO and \
                        (frame.width > frame.height) != (area_w > area_h):
                    frame = frame.rotate(90, expand=True)
                w, h = fit_rect(frame.width, frame.height, area_w, area_h, settings.scaling, img_dpi, dpi_x)
                x = mx + (area_w - w) // 2
                y = my + (area_h - h) // 2
                dc.StartPage()
                ImageWin.Dib(frame).draw(dc.GetHandleOutput(), (x, y, x + w, y + h))
                dc.EndPage()
        dc.EndDoc()
        started = False
        log(f"GDI 图片打印完成：{len(frames)} 页 x {settings.copies} 份")
        return len(frames)
    except PrintJobError:
        raise
    except Exception as e:  # noqa: BLE001
        raise PrintJobError(f"GDI 图片打印失败: {e}") from e
    finally:
        if started:
            try:
                dc.AbortDoc()
            except Exception:  # noqa: BLE001
                pass
        dc.DeleteDC()
