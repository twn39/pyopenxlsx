import pytest
from pyopenxlsx import (
    Workbook,
    Font,
    Fill,
    Alignment,
    Side,
    Border,
    XLLineStyle,
    XLPatternType,
    XLAlignmentStyle,
    load_workbook_async,
)
from pyopenxlsx.styles import is_date_format, _get_line_style
from pyopenxlsx._openxlsx import XLColor


@pytest.mark.asyncio
async def test_worksheet_async_methods(tmp_path):
    wb = Workbook()
    ws = wb.active

    # append_async
    await ws.append_async(["A", "B", "C"])
    assert ws["A1"].value == "A"

    # merge_cells_async / unmerge_cells_async
    await ws.merge_cells_async("A1:B1")
    assert "A1:B1" in ws.merges
    await ws.unmerge_cells_async("A1:B1")
    assert "A1:B1" not in ws.merges

    # set_cell_value_async
    await ws.set_cell_value_async(3, 3, "AsyncVal")
    assert ws.cell(3, 3).value == "AsyncVal"

    # protect_async / unprotect_async
    await ws.protect_async(password="123")
    assert ws.protection["protected"] is True
    await ws.unprotect_async()
    assert ws.protection["protected"] is False

    # get_rows_data_async / get_row_values_async / get_cell_value_async
    rows = await ws.get_rows_data_async()
    assert len(rows) >= 3
    row_vals = await ws.get_row_values_async(1)
    assert row_vals[0] == "A"
    cell_val = await ws.get_cell_value_async(3, 3)
    assert cell_val == "AsyncVal"

    # write_range_async: use a real numpy array
    import numpy as np

    data = np.array([[1.1, 2.2], [3.3, 4.4]], dtype=np.float64)
    await ws.write_range_async(10, 10, data)

    # get_range_values_async: read back into numpy
    res = await ws.get_range_values_async(10, 10, 11, 11)
    assert np.allclose(res, data)

    # write_rows_async / write_row_async
    await ws.write_rows_async(15, [["R15C1", "R15C2"], ["R16C1"]])
    await ws.write_row_async(17, ["R17C1", "R17C2"])
    assert ws.cell(15, 1).value == "R15C1"
    assert ws.cell(17, 2).value == "R17C2"

    # set_cells_async
    await ws.set_cells_async([(20, 1, "V1"), (20, 2, "V2")])
    assert ws.cell(20, 1).value == "V1"

    # save_async
    out = tmp_path / "async_save.xlsx"
    await wb.save_async(str(out))
    assert out.exists()

    # load_workbook_async
    wb2 = await load_workbook_async(str(out))
    assert wb2.active.title == ws.title

    # extract_images_async (even if empty)
    img_dir = tmp_path / "images"
    await wb2.extract_images_async(str(img_dir))

    wb.close()
    wb2.close()


def test_styles_more_coverage():
    # Font with direct XLColor in constructor
    color = XLColor("FF0000")
    f = Font(color=color)
    assert f.color() is color
    f.set_color("00FF00")  # string path
    assert f.color().hex().lower() == "ff00ff00"

    # Fill with XLPatternType and color object
    fill = Fill(pattern_type=XLPatternType.Gray125, color=color, background_color=color)
    assert fill.pattern_type() == XLPatternType.Gray125
    fill.set_color(color)
    fill.set_background_color(color)

    # Fill with string pattern type
    fill2 = Fill(pattern_type="darkGrid")
    assert fill2.pattern_type() == XLPatternType.DarkGrid

    # Alignment with XLAlignmentStyle
    align = Alignment(horizontal=XLAlignmentStyle.Center, vertical=XLAlignmentStyle.Top)
    assert align.horizontal() == XLAlignmentStyle.Center

    # Alignment with string mapping
    align2 = Alignment(horizontal="justify", vertical="distributed")
    assert align2.horizontal() == XLAlignmentStyle.Justify
    assert align2.vertical() == XLAlignmentStyle.Distributed

    # is_date_format complex strings
    assert is_date_format('[Red]"Date: "yyyy-mm-dd') is True
    assert is_date_format('0.00" seconds"') is False
    assert is_date_format(None) is False

    # _get_line_style unknown
    assert _get_line_style("unknown_style") == "unknown_style"

    # Side color object
    s = Side(style="thick", color=color)
    assert s.color() is color

    # Border setters
    b = Border()
    b.set_left(XLLineStyle.Thin, color)
    b.set_right(XLLineStyle.Thin, color)
    b.set_top(XLLineStyle.Thin, color)
    b.set_bottom(XLLineStyle.Thin, color)
    b.set_diagonal(XLLineStyle.Thin, color)
    assert b.left().style() == XLLineStyle.Thin


def test_worksheet_range_two_args():
    wb = Workbook()
    ws = wb.active
    # worksheet.py line 106: range with 2 args
    rng = ws.range("A1", "B2")
    assert rng is not None


def test_workbook_properties_more():
    wb = Workbook()
    # workbook.py line 190, 277 etc properties
    props = wb.properties
    _ = props["created"]  # access via __getitem__
    _ = props["modified"]
    wb.close()


def test_cell_line_108():
    wb = Workbook()
    ws = wb.active
    # cell.py line 108: cf.font_index()
    # Accessing cell.font triggers this
    ws["A1"].value = "Test"
    f = ws["A1"].font
    assert f is not None
    wb.close()
