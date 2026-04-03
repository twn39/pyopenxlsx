import pytest
from pyopenxlsx import Workbook

@pytest.mark.asyncio
async def test_range_extras():
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1).value = 1
    ws.cell(row=2, column=2).value = 2
    
    rng = ws.range("A1:B2")
    assert rng.address == "A1:B2"
    assert rng.num_rows == 2
    assert rng.num_columns == 2
    
    cells = list(rng)
    assert len(cells) == 4
    assert cells[0].value == 1
    assert cells[-1].value == 2
    
    rng.clear()
    assert ws.cell(row=1, column=1).value is None
    
    ws.cell(row=1, column=1).value = 1
    await rng.clear_async()
    assert ws.cell(row=1, column=1).value is None

def test_range_without_worksheet():
    from pyopenxlsx.range import Range
    from pyopenxlsx._openxlsx import XLDocument
    
    doc = XLDocument()
    doc.create("dummy.xlsx")
    wb = doc.workbook()
    ws = wb.worksheet("Sheet1")
    raw_range = ws.range("A1:A2")
    
    rng = Range(raw_range, None)
    cells = list(rng)
    assert len(cells) == 2

