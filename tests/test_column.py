import pytest
from pyopenxlsx import Workbook, load_workbook


def test_column_properties(tmp_path):
    wb = Workbook()
    ws = wb.active

    # Get column A (1)
    col_a = ws.column(1)
    col_a.width = 20.5
    col_a.hidden = True

    # Get column B ('B')
    col_b = ws.column("B")
    col_b.width = 15.0

    filename = tmp_path / "test_column.xlsx"
    wb.save(filename)

    # Reload and verify
    wb2 = load_workbook(str(filename))
    ws2 = wb2.active

    assert ws2.column(1).width == pytest.approx(20.5, 0.1)
    assert ws2.column(1).hidden is True
    assert ws2.column(2).width == pytest.approx(15.0, 0.1)
    assert ws2.column(2).hidden is False
    wb2.close()
    wb.close()


def test_column_style(tmp_path):
    wb = Workbook()
    ws = wb.active

    # Add a style for the column
    style_idx = wb.add_style(number_format="0.00%")

    col_c = ws.column("C")
    col_c.style_index = style_idx

    filename = tmp_path / "test_col_style.xlsx"
    wb.save(filename)

    wb2 = load_workbook(str(filename))
    ws2 = wb2.active
    assert ws2.column(3).style_index == style_idx
    wb2.close()
    wb.close()
