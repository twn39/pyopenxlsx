import os
from pyopenxlsx._openxlsx import XLDocument


def test_context_manager():
    """Verify context manager works and closes document."""
    doc_path = "test_ctx.xlsx"
    if os.path.exists(doc_path):
        os.remove(doc_path)

    with XLDocument() as doc:
        doc.create(doc_path)
        if not doc.workbook().sheet_exists("Sheet1"):
            doc.workbook().add_worksheet("Sheet1")
        doc.workbook().add_worksheet("Sheet2")
        # Ensure we can use it
        assert doc.workbook().sheet_exists("Sheet2")

    # Validating correct cleanup is hard without mocking or inspecting internal state,
    # but successful execution implies no double-free or crash.
    if os.path.exists(doc_path):
        os.remove(doc_path)


def test_keep_alive():
    """Verify child objects keep parent alive."""
    doc_path = "test_mem.xlsx"
    if os.path.exists(doc_path):
        os.remove(doc_path)

    def get_cell():
        doc = XLDocument()
        doc.create(doc_path)
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

    if os.path.exists(doc_path):
        os.remove(doc_path)
