import pytest
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLDocument


def test_range_basic(tmp_path):
    doc_path = tmp_path / "test_range.xlsx"
    doc = XLDocument()
    doc.create(str(doc_path))
    wb = doc.workbook()
    ws = wb.worksheet("Sheet1")

    rng = ws.range("A1:C3")
    # Test both raw and wrapped API if possible, but here we focus on wrapped
    assert rng.address() == "A1:C3"
    assert rng.num_rows() == 3
    assert rng.num_columns() == 3

    rng2 = ws.range("A1", "B2")
    assert rng2.address() == "A1:B2"
    assert rng2.num_rows() == 2
    assert rng2.num_columns() == 2

    doc.save()
    doc.close()


def test_range_iteration(tmp_path):
    doc_path = tmp_path / "test_iter.xlsx"
    doc = XLDocument()
    doc.create(str(doc_path))
    wb = doc.workbook()
    ws = wb.worksheet("Sheet1")

    # Set some values
    ws.cell("A1").value = 1
    ws.cell("A2").value = 2
    ws.cell("B1").value = 3
    ws.cell("B2").value = 4

    rng = ws.range("A1:B2")
    cells = list(rng)
    assert len(cells) == 4

    values = [cell.value for cell in cells]
    # Iteration order is row by row: A1, B1, A2, B2?
    # Actually, XLCellIterator in OpenXLSX usually iterates row by row, then column by column.
    # A1, B1, A2, B2
    assert 1 in values
    assert 2 in values
    assert 3 in values
    assert 4 in values

    # Check order
    assert cells[0].value == 1  # A1
    assert cells[1].value == 3  # B1
    assert cells[2].value == 2  # A2
    assert cells[3].value == 4  # B2

    doc.save()
    doc.close()


def test_range_clear(tmp_path):
    doc_path = tmp_path / "test_clear.xlsx"
    doc = XLDocument()
    doc.create(str(doc_path))
    wb = doc.workbook()
    ws = wb.worksheet("Sheet1")

    ws.cell("A1").value = "data"
    ws.cell("B2").value = 123

    rng = ws.range("A1:B2")
    rng.clear()

    assert ws.cell("A1").value is None
    assert ws.cell("B2").value is None

    doc.save()
    doc.close()


def test_wrapped_range_properties():
    wb = Workbook()
    ws = wb.active
    rng = ws.range("A1:B2")

    assert rng.address == "A1:B2"
    assert rng.num_rows == 2
    assert rng.num_columns == 2


@pytest.mark.asyncio
async def test_range_clear_async(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "async"
    rng = ws.range("A1:A1")

    await rng.clear_async()
    assert ws["A1"].value is None
    await wb.close_async()
