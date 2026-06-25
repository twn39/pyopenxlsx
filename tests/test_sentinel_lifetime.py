import pytest
from pyopenxlsx import Workbook


def test_closed_sentinel_protection():
    """Verify that operations on closed Workbook/Worksheet/Cell raise ValueError gracefully instead of segfaulting."""
    wb = Workbook()
    ws = wb.active
    cell = ws.cell(1, 1, "sentinel")

    # Verify everything works before closing
    assert cell.value == "sentinel"

    # Close workbook
    wb.close()

    # Now, accessing cell properties should raise ValueError
    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        _ = cell.value

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        cell.value = "new_val"

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        _ = cell.comment

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        cell.comment = "new_comment"

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        _ = cell.formula

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        cell.formula = "=SUM(A1:A5)"

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        _ = cell.style_index

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        cell.style_index = 2

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        _ = cell.is_date

    # Accessing worksheet operations should raise ValueError
    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        ws.cell(1, 2)

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        _ = ws["B2"]

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        ws.set_cell_value(2, 2, "val")

    with pytest.raises(ValueError, match="I/O operation on closed Workbook/Worksheet."):
        ws.write_rows(1, [[1, 2], [3, 4]])
