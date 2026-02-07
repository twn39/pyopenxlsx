import pytest
import asyncio
from pyopenxlsx import Workbook, Style, Font, Fill, Worksheet
from pyopenxlsx._openxlsx import XLProperty


def test_document_properties_deletion():
    wb = Workbook()
    props = wb.properties
    props.title = "Test Title"
    assert props.title == "Test Title"

    # Delete by XLProperty
    del props[XLProperty.Title]
    assert props.title == ""

    # Delete by string
    props.creator = "Test Creator"
    del props["creator"]
    assert props.creator == ""

    # Delete app property
    props["CustomProp"] = "Value"
    del props["CustomProp"]


def test_workbook_save_error():
    wb = Workbook()
    with pytest.raises(ValueError, match="No filename specified"):
        wb.save()


def test_workbook_context_manager():
    with Workbook() as wb:
        assert wb.active is not None


def test_add_style_with_style_object():
    wb = Workbook()
    s = Style(font=Font(bold=True), fill=Fill(color="FF0000"))
    idx = wb.add_style(s)
    assert idx >= 0


def test_workbook_active_setter():
    wb = Workbook()
    ws2 = wb.create_sheet("Sheet2")
    wb.active = ws2
    assert wb.active.title == "Sheet2"

    with pytest.raises(TypeError, match="Must be a Worksheet object"):
        wb.active = "Not a worksheet"


def test_workbook_create_sheet_index():
    wb = Workbook()
    ws = wb.create_sheet("First", index=0)
    assert ws.index == 0
    assert wb.sheetnames[0] == "First"


def test_workbook_copy_worksheet():
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Original"
    ws_copy = wb.copy_worksheet(ws)
    assert ws_copy.title == "Sheet1 Copy"
    # Test multiple copies
    ws_copy2 = wb.copy_worksheet(ws)
    assert ws_copy2.title == "Sheet1 Copy1"


def test_workbook_delitem():
    wb = Workbook()
    wb.create_sheet("ToDel")
    assert "ToDel" in wb.sheetnames
    del wb["ToDel"]
    assert "ToDel" not in wb.sheetnames

    with pytest.raises(KeyError):
        del wb["NonExistent"]


def test_worksheet_rows_iterator():
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = 1
    ws["B1"].value = 2
    ws["A2"].value = 3
    ws["B2"].value = 4

    rows = list(ws.rows)
    assert len(rows) == 2
    assert len(rows[0]) == 2
    assert rows[0][0].value == 1
    assert rows[1][1].value == 4


def test_worksheet_set_column_format_int():
    wb = Workbook()
    ws = wb.active
    idx = wb.add_style(font=Font(italic=True))
    # Now supported via int
    ws.set_column_format(1, idx)


def test_worksheet_add_image_errors(tmp_path):
    wb = Workbook()
    ws = wb.active

    # File not found
    with pytest.raises(FileNotFoundError):
        ws.add_image("non_existent.png")

    # Unsupported format
    bad_file = tmp_path / "test.txt"
    bad_file.write_text("not an image")
    with pytest.raises(ValueError, match="Unsupported image format"):
        ws.add_image(str(bad_file))


def test_worksheet_sheet_state():
    wb = Workbook()
    ws = wb.active
    wb.create_sheet("Other")  # Need at least one visible sheet
    assert ws.sheet_state == "visible"
    ws.sheet_state = "hidden"
    assert ws.sheet_state == "hidden"
    ws.sheet_state = "very_hidden"
    assert ws.sheet_state == "very_hidden"
    ws.sheet_state = "visible"
    assert ws.sheet_state == "visible"


def test_workbook_add_style_number_format_custom():
    wb = Workbook()
    # Test existing custom format
    fmt = "#,##0.00"
    idx1 = wb.add_style(number_format=fmt)
    idx2 = wb.add_style(number_format=fmt)
    # Cell format indices are always unique because add_style creates a new XF
    assert idx1 != idx2

    # Test new custom format
    idx3 = wb.add_style(number_format="0.000%")
    assert idx3 != idx1


def test_cell_style_properties():
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]

    # Test style properties (mostly getters)
    assert cell.font is not None
    assert cell.fill is not None
    assert cell.border is not None
    assert cell.alignment is not None
    # Cell doesn't have number_format property directly, but we can check style_index
    assert cell.style_index == 0


def test_cell_comment_properties():
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]

    assert cell.comment is None

    cell.comment = "Hello"
    assert cell.comment == "Hello"

    cell.comment = None
    assert cell.comment is None


def test_workbook_errors():
    wb = Workbook()

    # Test sheet access errors
    with pytest.raises(KeyError):
        _ = wb["NonExistent"]

    with pytest.raises(IndexError):
        _ = wb.sheetnames[99]

    with pytest.raises(KeyError):
        del wb["NonExistent"]


def test_cell_value_errors():
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]

    # Test unsupported value type
    with pytest.raises(TypeError):
        cell.value = {"a": 1}


def test_worksheet_errors_more():
    wb = Workbook()
    ws = wb.active

    # Test comment on cell without worksheet (should not happen with ws["A1"])
    from pyopenxlsx.cell import Cell

    c = Cell(ws["A1"]._cell, worksheet=None)
    assert c.comment is None
    with pytest.raises(ValueError):
        c.comment = "Fail"


def test_worksheet_more_coverage(tmp_path):
    wb = Workbook()
    ws = wb.active

    # Line 79: TypeError in __getitem__
    with pytest.raises(TypeError):
        _ = ws[123]

    # Line 108: TypeError in range
    with pytest.raises(TypeError):
        ws.range(1, 2, 3)

    # Line 98 in worksheet.py: Cache hit in _get_cached_cell
    # Triggered via Range
    c1 = ws.cell(1, 1)
    rng = ws.range("A1:A1")
    c2 = next(iter(rng))
    assert c1 is c2

    # Line 15 in range.py: Range without worksheet
    from pyopenxlsx.range import Range as PyRange
    from pyopenxlsx._openxlsx import XLDocument

    test_file = tmp_path / "test_doc.xlsx"
    doc = XLDocument()
    doc.create(str(test_file))
    raw_ws = doc.workbook().worksheet("Sheet1")
    raw_rng = raw_ws.range("A1:A1")
    rng_no_ws = PyRange(raw_rng, worksheet=None)
    c3 = next(iter(rng_no_ws))
    assert c3.value is None

    # Trigger line 40 in worksheet.py (sheet_state getter "visible")
    assert ws.sheet_state == "visible"

    # Line 244: .jpeg extension
    img_path = tmp_path / "test.jpeg"
    from PIL import Image

    Image.new("RGB", (10, 10)).save(img_path)
    ws.add_image(str(img_path), width=10, height=10)

    # Line 250-260: Auto detect width/height
    ws.add_image(str(img_path))  # Should use PIL to detect

    # Line 276: add_image_async
    async def test_async():
        await ws.add_image_async(str(img_path), anchor="B2")

    asyncio.run(test_async())


def test_cell_value_date_error():
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]
    # Trigger line 45 in cell.py (serial_to_datetime failure)
    # We need is_date to be True but value to be invalid for serial_to_datetime
    idx = wb.add_style(number_format="yyyy-mm-dd")
    cell.style_index = idx
    cell.value = 1e15  # Overflow serial
    assert cell.value == 1e15  # Should catch exception and return raw val


def test_cell_no_workbook(tmp_path):
    from pyopenxlsx.cell import Cell
    from pyopenxlsx._openxlsx import XLDocument

    test_file = tmp_path / "test_cell.xlsx"
    doc = XLDocument()
    doc.create(str(test_file))
    raw_ws = doc.workbook().worksheet("Sheet1")
    raw_cell = raw_ws.cell(1, 1)

    c = Cell(raw_cell, worksheet=None)
    assert c.font is None
    assert c.fill is None
    assert c.border is None
    assert c.alignment is None
    assert c.is_date is False


def test_workbook_more_coverage_2():
    wb = Workbook()
    # Line 91, 95: last_modified_by
    wb.properties.last_modified_by = "Tester"
    assert wb.properties.last_modified_by == "Tester"

    # Line 211, 219, 227: add_style with int
    idx = wb.add_style(font=0, fill=0, border=0)
    assert idx >= 0

    # Line 246: add_style with int number_format
    idx2 = wb.add_style(number_format=14)
    assert idx2 >= 0

    # Line 335-338: create_sheet default title
    ws1 = wb.create_sheet()
    assert ws1.title == "Sheet2"
    ws2 = wb.create_sheet()
    assert ws2.title == "Sheet3"

    # Line 324: active fallback
    wb.workbook.clear_active_tab()
    assert wb.active.title == "Sheet1"


def test_styles_border_direct_strings():
    from pyopenxlsx import Border, Side

    # Test Border with strings instead of Side objects
    b = Border(top="thin", bottom="thick", diagonal="dashed")
    # This triggers the 'else' blocks in Border.__init__
    assert b is not None

    # Test Border with Side objects
    s = Side(style="thin", color="FF0000")
    b2 = Border(left=s, right=s, top=s, bottom=s, diagonal=s)
    assert b2 is not None


def test_is_date_format_int():
    from pyopenxlsx.styles import is_date_format

    assert is_date_format(14) is True
    assert is_date_format(1) is False


def test_cell_standard_date_format():
    wb = Workbook()
    ws = wb.active
    idx = wb.add_style(number_format=14)  # Standard date
    cell = ws["A1"]
    cell.style_index = idx
    assert cell.is_date is True


def test_active_property_simple():
    wb = Workbook()
    # Should hit line 319 in workbook.py
    assert wb.active is not None
    assert isinstance(wb.active, Worksheet)


def test_add_image_no_pillow(tmp_path, monkeypatch):
    import sys

    # Create image using real PIL before mocking
    from PIL import Image

    img_path = tmp_path / "test_no_pil.png"
    Image.new("RGB", (10, 10)).save(img_path)

    # Mock ImportError for PIL
    with monkeypatch.context() as m:
        m.setitem(sys.modules, "PIL", None)
        m.setitem(sys.modules, "PIL.Image", None)

        wb = Workbook()
        ws = wb.active

        with pytest.raises(ImportError, match="Pillow is required"):
            ws.add_image(str(img_path))
