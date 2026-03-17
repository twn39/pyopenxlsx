import pytest
from pyopenxlsx import Workbook, XLPageOrientation


def test_page_margins(tmp_path):
    path = tmp_path / "test_margins.xlsx"
    wb = Workbook()
    ws = wb.active

    margins = ws.page_margins
    margins.left = 0.5
    margins.right = 0.5
    margins.top = 1.0
    margins.bottom = 1.0
    margins.header = 0.3
    margins.footer = 0.3

    assert margins.left == 0.5
    assert margins.top == 1.0

    wb.save(path)

    # Reload and verify
    wb2 = Workbook(path)
    ws2 = wb2.active
    margins2 = ws2.page_margins
    assert margins2.left == 0.5
    assert margins2.right == 0.5
    assert margins2.top == 1.0
    assert margins2.bottom == 1.0


def test_print_options(tmp_path):
    path = tmp_path / "test_print_options.xlsx"
    wb = Workbook()
    ws = wb.active

    options = ws.print_options
    options.grid_lines = True
    options.headings = True
    options.horizontal_centered = True
    options.vertical_centered = False

    assert options.grid_lines is True
    assert options.horizontal_centered is True

    wb.save(path)

    # Reload and verify
    wb2 = Workbook(path)
    ws2 = wb2.active
    options2 = ws2.print_options
    assert options2.grid_lines is True
    assert options2.headings is True
    assert options2.horizontal_centered is True
    assert options2.vertical_centered is False


def test_page_setup(tmp_path):
    path = tmp_path / "test_page_setup.xlsx"
    wb = Workbook()
    ws = wb.active

    setup = ws.page_setup
    setup.orientation = XLPageOrientation.Landscape
    setup.paper_size = 9 # A4
    setup.scale = 80
    setup.black_and_white = True

    assert setup.orientation == XLPageOrientation.Landscape
    assert setup.paper_size == 9
    assert setup.scale == 80

    wb.save(path)

    # Reload and verify
    wb2 = Workbook(path)
    ws2 = wb2.active
    setup2 = ws2.page_setup
    assert setup2.orientation == XLPageOrientation.Landscape
    assert setup2.paper_size == 9
    assert setup2.scale == 80
    assert setup2.black_and_white is True
