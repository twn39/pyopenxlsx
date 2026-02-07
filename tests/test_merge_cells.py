import pytest
import os
from pyopenxlsx import Workbook

TEST_FILE = "test_merges_v2.xlsx"


@pytest.fixture
def wb():
    if os.path.exists(TEST_FILE):
        os.remove(TEST_FILE)
    wb = Workbook()
    yield wb
    wb.close()  # Ensure handles are closed
    if os.path.exists(TEST_FILE):
        os.remove(TEST_FILE)


def test_merge_cells_api(wb):
    ws = wb.active

    # Check initial state
    assert len(ws.merges) == 0

    # Add merge
    ws.merges.append("A1:B2")
    assert len(ws.merges) == 1
    assert "A1:B2" in ws.merges
    assert ws.merges[0] == "A1:B2"

    # Find merge
    idx = ws.merges.find("A1:B2")
    assert idx != -1

    # Add another
    ws.merges.append("C3:D4")
    assert len(ws.merges) == 2
    assert "C3:D4" in ws.merges

    # Iteration
    merges_list = list(ws.merges)
    assert len(merges_list) == 2
    assert "A1:B2" in merges_list
    assert "C3:D4" in merges_list

    # Delete
    # Note: index might change after deletion if not careful,
    # but here we delete by found index.
    ws.merges.delete(idx)
    assert len(ws.merges) == 1
    assert "A1:B2" not in ws.merges
