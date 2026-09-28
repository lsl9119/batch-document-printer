from pathlib import Path

import pytest

from src.core.models import (Document, DuplexMode, FileType, Orientation, PrintSettings, PrintStatus,
                             detect_file_type, extensions_for_types, OfficeBackendPreference)

REQUIRED = {
    ".pdf": FileType.PDF, ".doc": FileType.WORD, ".docx": FileType.WORD, ".wps": FileType.WORD,
    ".xls": FileType.EXCEL, ".xlsx": FileType.EXCEL, ".et": FileType.EXCEL,
    ".ppt": FileType.PPT, ".pptx": FileType.PPT, ".dps": FileType.PPT,
    ".jpg": FileType.IMAGE, ".jpeg": FileType.IMAGE, ".png": FileType.IMAGE, ".bmp": FileType.IMAGE,
    ".tiff": FileType.IMAGE, ".txt": FileType.TEXT,
}


@pytest.mark.parametrize("ext,expected", sorted(REQUIRED.items()))
def test_detect_file_type_required_formats(ext, expected):
    assert detect_file_type(Path(f"a{ext}")) == expected
    assert detect_file_type(Path(f"A{ext.upper()}")) == expected


def test_unsupported_extension():
    with pytest.raises(ValueError):
        detect_file_type(Path("a.exe"))


def test_document_chinese_path_with_spaces(tmp_path):
    path = tmp_path / "中文 目录" / "季度 报告 (最终).docx"
    path.parent.mkdir()
    path.write_bytes(b"x" * 10)
    doc = Document(file_path=path)
    assert doc.file_type == FileType.WORD
    assert doc.file_name == "季度 报告 (最终).docx"
    assert doc.print_status == PrintStatus.PENDING
    assert doc.status_display == "等待"


def test_extensions_for_types():
    exts = extensions_for_types({"excel": True})
    assert exts == {".xls", ".xlsx", ".et"}
    assert ".dps" in extensions_for_types(None)


def test_print_settings_duplex_modes():
    s = PrintSettings(duplex=False)
    assert s.effective_duplex == DuplexMode.SIMPLEX and s.duplex_mode_str == "simplex"
    s = PrintSettings(duplex=True, duplex_mode=DuplexMode.DUPLEX_SHORT)
    assert s.duplex_mode_str == "duplexshort" and s.duplex_display == "双面-短边翻转"
    s = PrintSettings(duplex=True, duplex_mode=DuplexMode.SIMPLEX)
    assert s.effective_duplex == DuplexMode.SIMPLEX


def test_print_settings_roundtrip_and_legacy():
    s = PrintSettings(printer_name="P", paper_size="A3", copies=3, duplex=True,
                      duplex_mode=DuplexMode.DUPLEX_SHORT, orientation=Orientation.LANDSCAPE,
                      office_force_page_setup=True, office_backend=OfficeBackendPreference.WPS)
    assert PrintSettings.from_dict(s.to_dict()) == s
    # v5.x 配置: duplex_mode="duplex"，没有 orientation auto / office 字段
    legacy = PrintSettings.from_dict({"printer_name": "P", "duplex": True, "duplex_mode": "duplex",
                                      "orientation": "portrait", "copies": "2"})
    assert legacy.effective_duplex == DuplexMode.DUPLEX_LONG
    assert legacy.copies == 2
    assert legacy.office_backend == OfficeBackendPreference.AUTO
    # 损坏的值回落为默认值
    broken = PrintSettings.from_dict({"orientation": "diagonal", "copies": "abc", "color_mode": "x"})
    assert broken.orientation == Orientation.AUTO and broken.copies == 1


def test_copies_minimum_one():
    assert PrintSettings(copies=0).copies == 1
