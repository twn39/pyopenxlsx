import os
from pyopenxlsx._openxlsx import XLDocument


def test_context_manager(tmp_path):
    """Verify context manager works and closes document."""
    doc_path = tmp_path / "test_ctx.xlsx"

    with XLDocument() as doc:
        doc.create(str(doc_path))
        if not doc.workbook().sheet_exists("Sheet1"):
            doc.workbook().add_worksheet("Sheet1")
        doc.workbook().add_worksheet("Sheet2")
        # Ensure we can use it
        assert doc.workbook().sheet_exists("Sheet2")


def test_keep_alive(tmp_path):
    """Verify child objects keep parent alive."""
    doc_path = tmp_path / "test_mem.xlsx"

    def get_cell():
        doc = XLDocument()
        doc.create(str(doc_path))
        wb = doc.workbook()
        ws = wb.worksheet("Sheet1")
        return ws.cell("A1")

    # doc goes out of scope here.
    # If keep_alive works, cell should still be valid.
    # In C++, if doc was destroyed, accessing cell would segfault or read garbage.
    # With keep_alive, doc stays alive until cell is GC'd.
    cell = get_cell()
    cell.value = "Alive"
    assert cell.value == "Alive"

    # Explicitly delete cell to release keep-alive reference to doc
    del cell
