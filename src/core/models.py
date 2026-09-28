"""
数据模型定义
定义了应用中使用的所有数据结构
"""
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional


class FileType(Enum):
    """文件类型枚举"""
    WORD = "word"
    PPT = "ppt"
    EXCEL = "excel"
    PDF = "pdf"
    IMAGE = "image"
    TEXT = "text"


# 扩展名 → 文件类型（单一来源，DocumentManager / 处理器 / GUI 过滤器共用）
EXTENSION_TYPE_MAP: Dict[str, FileType] = {
    '.doc': FileType.WORD, '.docx': FileType.WORD, '.wps': FileType.WORD,
    '.ppt': FileType.PPT, '.pptx': FileType.PPT, '.dps': FileType.PPT,
    '.xls': FileType.EXCEL, '.xlsx': FileType.EXCEL, '.et': FileType.EXCEL,
    '.pdf': FileType.PDF,
    '.jpg': FileType.IMAGE, '.jpeg': FileType.IMAGE, '.png': FileType.IMAGE,
    '.bmp': FileType.IMAGE, '.tiff': FileType.IMAGE, '.tif': FileType.IMAGE,
    '.webp': FileType.IMAGE,
    '.txt': FileType.TEXT,
}

# 需要 Office 组件（WPS / Microsoft Office）才能打印的文件类型
OFFICE_FILE_TYPES = {FileType.WORD, FileType.EXCEL, FileType.PPT}


def extensions_for_types(enabled_file_types: Optional[Dict[str, bool]] = None) -> set:
    """根据启用的类型字典返回允许的扩展名集合；None 表示全部"""
    if enabled_file_types is None:
        return set(EXTENSION_TYPE_MAP)
    return {ext for ext, ftype in EXTENSION_TYPE_MAP.items()
            if enabled_file_types.get(ftype.value, False)}


def detect_file_type(file_path: Path) -> FileType:
    """根据扩展名识别文件类型，不支持时抛出 ValueError"""
    suffix = Path(file_path).suffix.lower()
    try:
        return EXTENSION_TYPE_MAP[suffix]
    except KeyError:
        raise ValueError(f"不支持的文件类型: {suffix}") from None


class PrintStatus(Enum):
    """打印状态枚举"""
    PENDING = "pending"      # 等待
    PRINTING = "printing"    # 正在打印
    COMPLETED = "completed"  # 成功
    ERROR = "error"          # 失败
    CANCELLED = "cancelled"  # 已取消


PRINT_STATUS_TEXT = {
    PrintStatus.PENDING: "等待",
    PrintStatus.PRINTING: "正在打印",
    PrintStatus.COMPLETED: "成功",
    PrintStatus.ERROR: "失败",
    PrintStatus.CANCELLED: "已取消",
}


class ColorMode(Enum):
    """颜色模式枚举"""
    COLOR = "color"         # 彩色
    GRAYSCALE = "grayscale" # 黑白


class Orientation(Enum):
    """页面方向枚举"""
    AUTO = "auto"           # 自动（跟随文档/驱动）
    PORTRAIT = "portrait"   # 纵向
    LANDSCAPE = "landscape" # 横向


class DuplexMode(Enum):
    """双面打印模式枚举"""
    SIMPLEX = "simplex"           # 单面
    DUPLEX_LONG = "duplexlong"    # 双面-长边翻转
    DUPLEX_SHORT = "duplexshort"  # 双面-短边翻转


class ScalingMode(Enum):
    """缩放模式枚举（PDF/图片）"""
    FIT = "fit"           # 适合页面
    SHRINK = "shrink"     # 仅缩小
    NOSCALE = "noscale"   # 无缩放


class OfficeBackendPreference(Enum):
    """Office 引擎偏好"""
    AUTO = "auto"         # 自动：优先 WPS，缺失组件时回退 Microsoft Office
    WPS = "wps"           # 仅 WPS Office
    MS_OFFICE = "msoffice"  # 仅 Microsoft Office


@dataclass
class Document:
    """文档数据模型"""
    file_path: Path
    file_name: str = field(init=False)
    file_type: FileType = field(init=False)
    file_size: int = field(init=False)
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    added_time: datetime = field(default_factory=datetime.now)
    print_status: PrintStatus = PrintStatus.PENDING
    last_error: str = ""
    last_backend: str = ""

    def __post_init__(self):
        """初始化后处理"""
        self.file_path = Path(self.file_path)
        self.file_name = self.file_path.name
        self.file_size = self.file_path.stat().st_size if self.file_path.exists() else 0
        self.file_type = detect_file_type(self.file_path)

    @property
    def size_mb(self) -> float:
        """获取文件大小（MB）"""
        return round(self.file_size / (1024 * 1024), 2)

    @property
    def type_display(self) -> str:
        """获取文件类型显示名称"""
        type_map = {
            FileType.WORD: "文字文档",
            FileType.PPT: "演示文稿",
            FileType.EXCEL: "表格",
            FileType.PDF: "PDF文件",
            FileType.IMAGE: "图片文件",
            FileType.TEXT: "文本文件"
        }
        return type_map.get(self.file_type, "未知")

    @property
    def status_display(self) -> str:
        return PRINT_STATUS_TEXT.get(self.print_status, "未知")


@dataclass
class PrintSettings:
    """打印设置数据模型"""
    printer_name: str = ""
    paper_size: str = "A4"
    copies: int = 1

    duplex: bool = False                               # 是否双面（兼容旧配置）
    duplex_mode: DuplexMode = DuplexMode.DUPLEX_LONG   # 双面时的翻转方式

    color_mode: ColorMode = ColorMode.GRAYSCALE
    orientation: Orientation = Orientation.AUTO
    scaling: ScalingMode = ScalingMode.FIT

    # Office 文档是否强制套用上面的纸张/方向（默认遵循文档自身页面设置）
    office_force_page_setup: bool = False
    # Office 引擎偏好
    office_backend: OfficeBackendPreference = OfficeBackendPreference.AUTO

    def __post_init__(self):
        try:
            self.copies = max(1, int(self.copies))
        except (TypeError, ValueError):
            self.copies = 1
        # duplex_mode 为 SIMPLEX 时等价于关闭双面
        if self.duplex_mode == DuplexMode.SIMPLEX:
            self.duplex = False
            self.duplex_mode = DuplexMode.DUPLEX_LONG

    # 便捷属性
    @property
    def color(self) -> bool:
        """是否彩色打印"""
        return self.color_mode == ColorMode.COLOR

    @property
    def effective_duplex(self) -> DuplexMode:
        """实际双面模式（未开启双面时为 SIMPLEX）"""
        return self.duplex_mode if self.duplex else DuplexMode.SIMPLEX

    @property
    def scaling_str(self) -> str:
        return self.scaling.value

    @property
    def orientation_str(self) -> str:
        return self.orientation.value

    @property
    def duplex_mode_str(self) -> str:
        """SumatraPDF 使用的双面参数字符串"""
        return self.effective_duplex.value

    @property
    def duplex_display(self) -> str:
        return {
            DuplexMode.SIMPLEX: "单面",
            DuplexMode.DUPLEX_LONG: "双面-长边翻转",
            DuplexMode.DUPLEX_SHORT: "双面-短边翻转",
        }[self.effective_duplex]

    @property
    def orientation_display(self) -> str:
        return {Orientation.AUTO: "自动", Orientation.PORTRAIT: "纵向",
                Orientation.LANDSCAPE: "横向"}[self.orientation]

    @property
    def color_display(self) -> str:
        return "彩色" if self.color else "黑白"

    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'printer_name': self.printer_name,
            'paper_size': self.paper_size,
            'copies': self.copies,
            'duplex': self.duplex,
            'duplex_mode': self.duplex_mode.value,
            'color_mode': self.color_mode.value,
            'orientation': self.orientation.value,
            'scaling': self.scaling.value,
            'office_force_page_setup': self.office_force_page_setup,
            'office_backend': self.office_backend.value,
        }

    @staticmethod
    def _enum(enum_cls, value, default):
        try:
            return enum_cls(value)
        except ValueError:
            return default

    @classmethod
    def from_dict(cls, data: dict) -> 'PrintSettings':
        """从字典创建实例（兼容 v5.x 旧配置）"""
        data = data or {}
        raw_duplex_mode = data.get('duplex_mode', 'duplexlong')
        if raw_duplex_mode == 'duplex':  # v5.x 的 "duplex" 等同长边翻转
            raw_duplex_mode = 'duplexlong'
        return cls(
            printer_name=data.get('printer_name', '') or '',
            paper_size=data.get('paper_size', 'A4') or 'A4',
            copies=data.get('copies', 1),
            duplex=bool(data.get('duplex', False)),
            duplex_mode=cls._enum(DuplexMode, raw_duplex_mode, DuplexMode.DUPLEX_LONG),
            color_mode=cls._enum(ColorMode, data.get('color_mode', 'grayscale'), ColorMode.GRAYSCALE),
            orientation=cls._enum(Orientation, data.get('orientation', 'auto'), Orientation.AUTO),
            scaling=cls._enum(ScalingMode, data.get('scaling', 'fit'), ScalingMode.FIT),
            office_force_page_setup=bool(data.get('office_force_page_setup', False)),
            office_backend=cls._enum(OfficeBackendPreference, data.get('office_backend', 'auto'),
                                     OfficeBackendPreference.AUTO),
        )


def _default_enabled_types() -> dict:
    return {'word': True, 'ppt': True, 'excel': True, 'pdf': True, 'image': True, 'text': True}


@dataclass
class AppConfig:
    """应用配置数据模型"""
    last_printer: str = ""
    default_settings: PrintSettings = field(default_factory=PrintSettings)
    window_geometry: dict = field(default_factory=dict)
    recent_folders: List[str] = field(default_factory=list)
    # 文件类型过滤器设置（WPS 版默认全部启用，表格也可直接打印）
    enabled_file_types: dict = field(default_factory=_default_enabled_types)

    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'last_printer': self.last_printer,
            'default_settings': self.default_settings.to_dict(),
            'window_geometry': self.window_geometry,
            'recent_folders': self.recent_folders,
            'enabled_file_types': self.enabled_file_types
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'AppConfig':
        """从字典创建实例"""
        data = data or {}
        enabled_types = dict(data.get('enabled_file_types') or {})
        for file_type, default in _default_enabled_types().items():
            enabled_types.setdefault(file_type, default)

        return cls(
            last_printer=data.get('last_printer', ''),
            default_settings=PrintSettings.from_dict(data.get('default_settings', {})),
            window_geometry=data.get('window_geometry', {}) or {},
            recent_folders=data.get('recent_folders', []) or [],
            enabled_file_types=enabled_types
        )
