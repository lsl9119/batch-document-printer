import pytest

from src.backends import office_ops as ops
from src.core.models import Orientation, PrintSettings
from tests.fakes import FakeOfficeApp, password_error


def test_writer_print_named_args():
    app = FakeOfficeApp("KWPS.Application")
    doc = ops.open_writer_document(app, "C:/a b/中文.docx")
    method = ops.writer_print(app, doc, 3)
    assert "Background=False" in method
    assert doc.printouts[0]["kwargs"]["Copies"] == 3
    assert doc.printouts[0]["kwargs"]["Background"] is False
    ops.close_writer_document(doc)
    assert doc.closed and not app.open_docs


def test_writer_falls_back_when_named_args_unsupported():
    app = FakeOfficeApp("KWPS.Application", named_args_supported=False)
    doc = ops.open_writer_document(app, "a.wps")
    # 位置参数方式打开，带假密码
    assert app.open_calls[0][0][4] == ops.DUMMY_PASSWORD
    method = ops.writer_print(app, doc, 2)
    assert method == "PrintOut 位置参数"
    assert doc.printouts[0]["args"][7] == 2


def test_writer_password_protected_file_is_reported():
    app = FakeOfficeApp("KWPS.Application")
    app.open_error = password_error()
    with pytest.raises(ops.OfficeDocumentError) as info:
        ops.open_writer_document(app, "secret.docx")
    assert "密码保护" in str(info.value)
    assert len(app.open_calls) == 0  # 不会再尝试无密码打开（那样会弹出密码框卡住）


def test_writer_page_setup_only_when_forced():
    app = FakeOfficeApp("KWPS.Application")
    doc = ops.open_writer_document(app, "a.docx")
    settings = PrintSettings(paper_size="A3", orientation=Orientation.LANDSCAPE)
    assert ops.apply_writer_page_setup(doc, settings, lambda m: None) == []
    assert doc.PageSetup.Orientation == 1 and doc.PageSetup.PaperSize == 6


def test_spreadsheet_counts_visible_non_empty_sheets():
    app = FakeOfficeApp("KET.Application")
    app.sheet_specs = [dict(name="A4纵向", pages=2), dict(name="A3横向", pages=3),
                       dict(name="隐藏", visible=0, pages=9), dict(name="空表", has_content=False, pages=1),
                       dict(name="打印区域", pages=1, print_area="$A$1:$C$10")]
    wb = ops.open_workbook(app, "multi.xlsx")
    assert ops.spreadsheet_count_pages(wb) == 6
    method = ops.spreadsheet_print(app, wb, 2)
    assert "Workbook.PrintOut" in method and wb.printouts[0]["kwargs"]["Copies"] == 2


def test_spreadsheet_force_page_setup_applies_to_visible_sheets():
    app = FakeOfficeApp("KET.Application")
    app.sheet_specs = [dict(name="S1"), dict(name="S2", visible=0)]
    wb = ops.open_workbook(app, "a.et")
    ops.apply_spreadsheet_page_setup(wb, PrintSettings(paper_size="A3", orientation=Orientation.LANDSCAPE),
                                     lambda m: None)
    s1, s2 = wb.Worksheets.Item(1), wb.Worksheets.Item(2)
    assert (s1.PageSetup.Orientation, s1.PageSetup.PaperSize) == (2, 8)
    assert (s2.PageSetup.Orientation, s2.PageSetup.PaperSize) == (1, 9)


def test_spreadsheet_empty_workbook_fails_clearly():
    app = FakeOfficeApp("KET.Application")
    app.sheet_specs = [dict(name="空", has_content=False)]
    wb = ops.open_workbook(app, "empty.xlsx")
    with pytest.raises(ops.OfficeDocumentError, match="没有可打印的内容"):
        ops.spreadsheet_print(app, wb, 1)


def test_spreadsheet_worksheet_fallback_without_named_args():
    app = FakeOfficeApp("KET.Application", named_args_supported=False)
    app.sheet_specs = [dict(name="S1"), dict(name="S2")]
    wb = ops.open_workbook(app, "a.xls")
    method = ops.spreadsheet_print(app, wb, 1)
    assert method == "Workbook.PrintOut 位置参数"


def test_presentation_active_printer_readonly_is_tolerated():
    app = FakeOfficeApp("KWPP.Application", active_printer="Office Printer", active_printer_writable=False)
    pres = ops.open_presentation(app, "deck.dps")
    settings = PrintSettings(printer_name="Target Printer", copies=2)
    result = ops.presentation_print(pres, settings, lambda m: None)
    assert result["active_printer_writable"] is False
    assert "NumberOfCopies=2" in result["method"]
    assert app.printed[-1][3] == 2
    ops.close_presentation(pres)
    assert not app.open_docs


def test_presentation_active_printer_writable():
    app = FakeOfficeApp("KWPP.Application", active_printer="Office Printer")
    pres = ops.open_presentation(app, "deck.pptx")
    result = ops.presentation_print(pres, PrintSettings(printer_name="Target Printer"), lambda m: None)
    assert result["active_printer_writable"] is True
    assert pres.PrintOptions.ActivePrinter == "Target Printer"


def test_presentation_color_mode():
    app = FakeOfficeApp("KWPP.Application")
    pres = ops.open_presentation(app, "deck.pptx")
    ops.presentation_print(pres, PrintSettings(), lambda m: None)
    assert pres.PrintOptions.PrintColorType == ops.PP_PRINT_BLACK_AND_WHITE


def test_printer_matches_variants():
    assert ops.printer_matches("HP LaserJet on Ne01:", "HP LaserJet")
    assert ops.printer_matches("hp laserjet", "HP LaserJet")
    assert not ops.printer_matches("HP LaserJet 2", "HP LaserJet 200")
