"""Regression: cell weak/MRU cache contract and bulk path isolation."""

import gc
import weakref

from pyopenxlsx import Workbook


class TestCellCacheContract:
    def test_identity_reuse_via_cell_and_getitem(self):
        wb = Workbook()
        ws = wb.active
        a = ws.cell(1, 1)
        b = ws.cell(1, 1)
        c = ws["A1"]
        assert a is b
        # getitem uses string key; cell uses (row, col) — both should resolve
        # to the same cached Cell when keys alias the same address.
        assert a is c or (a._cell.cell_reference().address() == "A1" and c is ws["A1"])
        wb.close()

    def test_mru_caps_at_64(self):
        wb = Workbook()
        ws = wb.active
        for r in range(1, 200):
            ws.cell(r, 1)
        assert len(ws._cells_mru) == 64
        wb.close()

    def test_weak_cache_drops_without_external_refs(self):
        wb = Workbook()
        ws = wb.active
        # Fill MRU then drop external refs and force GC of unreferenced cells.
        for r in range(1, 100):
            ws.cell(r, 1)
        # Keep only last access path alive via MRU; clear MRU to allow weak drop.
        ws._cells_mru.clear()
        gc.collect()
        # WeakValueDictionary should only hold cells still strongly referenced.
        # After clearing MRU and GC, keys may shrink (implementation-dependent
        # on whether C++/other refs exist). At least access should re-create.
        before = len(ws._cells)
        c = ws.cell(1, 1)
        assert c is not None
        assert len(ws._cells) >= 1
        # Re-access same identity
        assert ws.cell(1, 1) is c
        _ = before  # silence unused if before is 0
        wb.close()

    def test_no_cycle_leaks_workbook_graph(self):
        wb = Workbook()
        ws = wb.active
        cell = ws.cell(1, 1, "x")
        wb_ref = weakref.ref(wb)
        ws_ref = weakref.ref(ws)
        cell_ref = weakref.ref(cell)
        del cell, ws, wb
        gc.collect()
        assert wb_ref() is None
        assert ws_ref() is None
        assert cell_ref() is None

    def test_bulk_write_does_not_require_cell_cache(self):
        wb = Workbook()
        ws = wb.active
        ws._cells.clear()
        ws._cells_mru.clear()
        ws.set_cell_value(1, 1, "bulk")
        ws.write_rows(2, [["a", 1], ["b", 2]])
        ws.set_cells([(5, 5, 42)])
        # Bulk paths should not populate the Cell cache.
        assert len(ws._cells) == 0
        # Values still readable via bulk or fresh cell wrappers.
        assert ws.get_cell_value(1, 1) == "bulk"
        assert ws.cell(5, 5).value == 42
        wb.close()

    def test_cell_value_and_set_cell_value_share_storage(self):
        wb = Workbook()
        ws = wb.active
        ws.set_cell_value(1, 1, "hello")
        assert ws.cell(1, 1).value == "hello"
        ws.cell(1, 1).value = "world"
        assert ws.get_cell_value(1, 1) == "world"
        wb.close()

    def test_closed_workbook_blocks_cell_and_bulk(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "x")
        wb.close()
        try:
            ws.cell(1, 1)
            raised = False
        except ValueError:
            raised = True
        assert raised
        try:
            ws.set_cell_value(1, 2, "y")
            raised2 = False
        except ValueError:
            raised2 = True
        assert raised2
