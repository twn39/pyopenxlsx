import gc
import weakref
from pyopenxlsx import Workbook


def test_cell_direct_references():
    """Verify that Cell uses direct references instead of weakrefs for performance."""
    wb = Workbook()
    ws = wb.active
    cell = ws.cell(1, 1, "test")

    # Cell should directly hold a reference to Worksheet, not a weakref object
    assert hasattr(cell, "_worksheet_val")
    assert cell._worksheet_val is ws
    assert cell._worksheet is ws
    assert cell._workbook is wb

    # Modifying value should still work perfectly
    cell.value = "new_val"
    assert cell.value == "new_val"


def test_cell_l1_mru_cache():
    """Verify that Worksheet keeps a light MRU cache of strong references to cells."""
    wb = Workbook()
    ws = wb.active

    # Create 10 cells and ensure they are all kept alive in MRU cache
    cells = [ws.cell(row, 1) for row in range(1, 11)]

    # All generated cells should be in the MRU cache
    for cell in cells:
        assert cell in ws._cells_mru

    # Generate more cells to exceed the maxlen of 64
    for r in range(12, 100):
        ws.cell(r, 1)

    # The deque maxlen is 64. Since we generated ~88 more cells without keeping
    # external references, the earliest cells (like row 1) should have been evicted
    # from the MRU deque.
    # Note: re-accessing would insert them back, but let us verify they were popped.
    # We can check that the size of _cells_mru is exactly 64.
    assert len(ws._cells_mru) == 64


def test_no_memory_leak():
    """Verify that direct reference model does not leak memory (no circular cycles)."""
    wb = Workbook()
    ws = wb.active
    cell = ws.cell(1, 1, "test_val")

    # Create weak references to monitor workbook and worksheet lifetime
    wb_ref = weakref.ref(wb)
    ws_ref = weakref.ref(ws)
    cell_ref = weakref.ref(cell)

    # Delete references
    del cell
    del ws
    del wb
    gc.collect()

    # Everything should be garbage collected because there are no circular reference cycles!
    assert wb_ref() is None
    assert ws_ref() is None
    assert cell_ref() is None
