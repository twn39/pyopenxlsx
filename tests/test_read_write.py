import os
from pyopenxlsx import Workbook, load_workbook


def test_create_and_write(tmp_path):
    """Test creating a new workbook and writing values."""
    filename = tmp_path / "test_write.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "TestSheet"

    # Write various types
    ws.cell(1, 1).value = "Hello"
    ws.cell(1, 2).value = 123
    ws.cell(1, 3).value = 3.14
    ws.cell(1, 4).value = True

    wb.save(filename)
    assert os.path.exists(filename)


def test_read_written_values(tmp_path):
    """Test reading back values that were written."""
    filename = tmp_path / "test_read.xlsx"
    wb = Workbook()
    ws = wb.active

    # Write data
    ws["A1"].value = "World"
    ws["B1"].value = 456
    ws["C1"].value = 1.23
    ws["D1"].value = False

    wb.save(filename)

    # Reload
    wb2 = load_workbook(str(filename))
    ws2 = wb2.active

    assert ws2["A1"].value == "World"
    assert ws2["B1"].value == 456
    assert abs(ws2["C1"].value - 1.23) < 1e-9
    assert ws2["D1"].value is False


def test_overwrite_values(tmp_path):
    """Test overwriting existing values."""
    filename = tmp_path / "test_overwrite.xlsx"
    wb = Workbook()
    ws = wb.active

    ws["A1"].value = "Original"
    wb.save(filename)

    # Reload and overwrite
    wb2 = load_workbook(str(filename))
    ws2 = wb2.active
    assert ws2["A1"].value == "Original"

    ws2["A1"].value = "New Value"
    wb2.save()  # Save to same file

    # Reload again
    wb3 = load_workbook(str(filename))
    ws3 = wb3.active
    assert ws3["A1"].value == "New Value"


def test_sheet_management(tmp_path):
    """Test creating and renaming sheets."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "First"

    wb.create_sheet("Second")
    wb.create_sheet("Third")

    assert wb.sheetnames == ["First", "Second", "Third"]

    del wb["Second"]
    assert wb.sheetnames == ["First", "Third"]

    filename = tmp_path / "test_sheets.xlsx"
    wb.save(filename)

    wb_loaded = load_workbook(str(filename))
    assert wb_loaded.sheetnames == ["First", "Third"]
