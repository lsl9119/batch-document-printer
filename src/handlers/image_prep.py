"""
打印前的图片预处理

手机拍摄的照片通常有数千万像素，并依赖 EXIF 方向标记显示为正向。直接交给 SumatraPDF
打印时，部分照片会出现“SumatraPDF 很快退出、返回成功，但打印机没有收到任务”的情况，
即使打出来，发送给打印机的数据量也很大。

这里在打印前：
- 按 EXIF 方向把图片转正（输出文件不再带 EXIF 方向，避免被重复旋转）
- 把超过纸张 300 dpi 所需像素的图片等比缩小（肉眼看不出差别，数据量小很多）
- 透明背景合成为白底，CMYK/16 位等模式转换为 RGB
- JPEG 照片输出 JPEG，其他格式输出无损 PNG

不需要处理的图片直接使用原文件；多页 TIFF 不处理。处理后的文件写入临时目录，打印完成后删除。
"""
import logging
import shutil
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Callable, Iterator, List, Optional

from ..core.models import PrintSettings, ScalingMode

logger = logging.getLogger("bdp.image_prep")

PRINT_DPI = 300
EXIF_ORIENTATION = 0x0112

# 纸张长边（毫米）
PAPER_LONG_SIDE_MM = {
    "A3": 420, "A4": 297, "A5": 210, "A6": 148, "B4": 364, "B5": 257,
    "LETTER": 279.4, "LEGAL": 355.6, "TABLOID": 431.8,
}
DEFAULT_LONG_SIDE_MM = 297


def _transpose_methods():
    from PIL import Image  # type: ignore
    t = Image.Transpose
    # 与 PIL.ImageOps.exif_transpose 相同的 EXIF 方向 → 变换对应关系
    return {2: t.FLIP_LEFT_RIGHT, 3: t.ROTATE_180, 4: t.FLIP_TOP_BOTTOM, 5: t.TRANSPOSE,
            6: t.ROTATE_270, 7: t.TRANSVERSE, 8: t.ROTATE_90}


try:
    TRANSPOSE_METHODS = _transpose_methods()
except ImportError:  # pragma: no cover - 未安装 Pillow 时图片处理器本身也不可用
    TRANSPOSE_METHODS = {}


def max_long_side_px(paper_size: str, dpi: int = PRINT_DPI) -> int:
    """在指定纸张上以 dpi 打印时，图片长边需要的最大像素数"""
    mm = PAPER_LONG_SIDE_MM.get((paper_size or "").strip().upper(), DEFAULT_LONG_SIDE_MM)
    return int(round(mm / 25.4 * dpi))


def _to_printable_mode(img):
    from PIL import Image  # type: ignore
    if img.mode in ("RGB", "L"):
        return img
    if img.mode == "P" and "transparency" in img.info:
        img = img.convert("RGBA")
    if img.mode in ("RGBA", "LA", "PA"):
        rgba = img.convert("RGBA")
        background = Image.new("RGB", rgba.size, (255, 255, 255))
        background.paste(rgba, mask=rgba.getchannel("A"))
        return background
    if img.mode in ("I;16", "I;16B", "I;16L", "I", "F"):
        return img.convert("L")
    return img.convert("RGB")


@contextmanager
def prepared_image(file_path: Path, settings: PrintSettings,
                   log: Optional[Callable[[str], None]] = None,
                   notes: Optional[List[str]] = None) -> Iterator[Path]:
    """
    with prepared_image(path, settings, log, notes) as print_path:
        ... 打印 print_path ...

    需要预处理时 print_path 为临时文件，否则为原文件。预处理失败时退回原文件。
    """
    log = log or (lambda m: logger.info(m))
    temp_dir = None
    print_path = file_path
    try:
        try:
            result = _prepare(file_path, settings, log)
        except Exception as e:  # noqa: BLE001 - 预处理失败不影响打印原文件
            log(f"图片预处理失败，直接打印原文件: {e}")
            result = None
        if result is not None:
            temp_dir, print_path, description = result
            log(f"图片预处理: {description}")
            if notes is not None:
                notes.append(description)
        yield print_path
    finally:
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)


def _prepare(file_path: Path, settings: PrintSettings, log):
    from PIL import Image  # type: ignore
    with Image.open(file_path) as img:
        # 多页 TIFF / 动图交给打印程序原样处理；手机照片常被识别为 MPO（多帧 JPEG），只取主图
        if getattr(img, "n_frames", 1) > 1 and img.format != "MPO":
            return None
        orientation = 1
        try:
            orientation = int(img.getexif().get(EXIF_ORIENTATION, 1) or 1)
        except Exception:  # noqa: BLE001 - EXIF 损坏时按无方向处理
            pass
        limit = max_long_side_px(settings.paper_size)
        too_large = settings.scaling != ScalingMode.NOSCALE and max(img.size) > limit
        rotate = orientation in TRANSPOSE_METHODS
        odd_mode = img.mode not in ("RGB", "L")
        if not (too_large or rotate or odd_mode):
            return None
        original_size = img.size
        original_mode = img.mode
        is_jpeg = img.format in ("JPEG", "MPO")
        dpi = img.info.get("dpi")
        if too_large and is_jpeg:
            # JPEG 按缩小后的尺寸解码，避免把数千万像素完整载入内存
            scale = limit / max(img.size)
            img.draft(None, (int(img.width * scale), int(img.height * scale)))
        work = img.copy()

    actions = []
    if rotate:
        actions.append("按拍摄方向转正")
    if odd_mode:
        actions.append(f"颜色模式 {original_mode} 转换")
    work = _to_printable_mode(work)
    if too_large:
        work.thumbnail((limit, limit), Image.LANCZOS)
        actions.append(f"缩小 {original_size[0]}x{original_size[1]} → {work.width}x{work.height}")
        dpi = (PRINT_DPI, PRINT_DPI)
    if rotate:
        work = work.transpose(TRANSPOSE_METHODS[orientation])

    temp_dir = Path(tempfile.mkdtemp(prefix="bdp_img_"))
    try:
        # 照片保持 JPEG；扫描件/截图等其他格式用无损 PNG，避免文字边缘出现压缩痕迹
        if is_jpeg:
            target = temp_dir / f"{file_path.stem}.jpg"
            save_kwargs = {"format": "JPEG", "quality": 92}
        else:
            target = temp_dir / f"{file_path.stem}.png"
            save_kwargs = {"format": "PNG"}
        if dpi:
            save_kwargs["dpi"] = tuple(float(v) for v in dpi[:2])
        work.save(target, **save_kwargs)
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise
    return temp_dir, target, "，".join(actions)
