import pytest
from pyopenxlsx import Workbook


def test_cell_access():
    wb = Workbook()
    ws = wb.active

    cell_a1 = ws["A1"]
    assert cell_a1.value is None

    ws["B2"].value = 42
    assert ws["B2"].value == 42

    ws["C3"].value = 3.14
    assert ws["C3"].value == 3.14

    ws["D4"].value = "OpenXLSX"
    assert ws["D4"].value == "OpenXLSX"

    ws["E5"].value = True
    assert ws["E5"].value is True


def test_cell_coordinates():
    wb = Workbook()
    ws = wb.active

    # Access by row/col (1-based index)
    ws.cell(row=1, column=1).value = "R1C1"
    assert ws["A1"].value == "R1C1"

    ws.cell(row=5, column=5).value = "R5C5"
    assert ws["E5"].value == "R5C5"


def test_unsupported_type():
    wb = Workbook()
    ws = wb.active
    with pytest.raises(TypeError):
        ws["A1"].value = [1, 2, 3]  # Lists not supported


def test_cell_style_properties():
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]

    # Create valid styles
    from pyopenxlsx import Font, Fill

    s1 = wb.add_style(font=Font(bold=True))
    s2 = wb.add_style(fill=Fill(color="FFFF00"))

    # Test style index
    cell.style_index = s1
    assert cell.style_index == s1
    assert cell.style == s1

    cell.style = s2
    assert cell.style_index == s2

    # Test font, fill, border, alignment (requires workbook)
    assert cell.font is not None
    assert cell.fill is not None
    assert cell.border is not None
    assert cell.alignment is not None


def test_cell_date_detection():
    import datetime

    wb = Workbook()
    ws = wb.active

    # Test date value setting and getting
    dt = datetime.datetime(2023, 1, 1, 12, 0, 0)
    ws["A1"].value = dt

    # We need to set a date format for it to be recognized as date on read
    style_date = wb.add_style(number_format="yyyy-mm-dd")
    ws["A1"].style_index = style_date

    assert ws["A1"].is_date is True
    assert isinstance(ws["A1"].value, datetime.datetime)

    # Test cache
    assert ws["A1"].is_date is True

    # Test invalid style index for is_date
    ws["A1"].style_index = 999  # Should return False and cache it
    assert ws["A1"].is_date is False


def test_cell_comment_errors():
    from pyopenxlsx.cell import Cell

    # Create a cell without a worksheet
    class MockRawCell:
        pass

    cell = Cell(MockRawCell(), worksheet=None)
    assert cell.comment is None

    with pytest.raises(
        ValueError, match="Cell must be associated with a worksheet to set comments"
    ):
        cell.comment = "Test"
