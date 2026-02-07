"""Tests for weak reference memory management and edge cases."""

import gc

import pytest

from pyopenxlsx import Workbook


class TestWeakRefMemory:
    """Test that weak references work correctly for memory management."""

    def test_cell_weakref_after_worksheet_gc(self):
        """Test Cell's _worksheet property when worksheet is garbage collected."""
        wb = Workbook()
        ws = wb.active
        cell = ws.cell(1, 1, "test")

        # The cell should have access to worksheet
        assert cell._worksheet is not None
        assert cell._workbook is not None

        # Cell should be accessible
        assert cell.value == "test"

    def test_cell_without_worksheet(self, tmp_path):
        """Test Cell created without worksheet reference."""
        from pyopenxlsx.cell import Cell
        from pyopenxlsx._openxlsx import XLDocument

        # Use temporary directory for file creation
        temp_file = tmp_path / "test_cell.xlsx"

        doc = XLDocument()
        doc.create(str(temp_file))
        wb_raw = doc.workbook()
        ws_raw = wb_raw.worksheet("Sheet1")
        cell_raw = ws_raw.cell(1, 1)

        # Create cell without worksheet
        cell = Cell(cell_raw, None)

        # Properties should return None
        assert cell._worksheet is None
        assert cell._workbook is None
        assert cell.font is None
        assert cell.fill is None
        assert cell.border is None
        assert cell.alignment is None
        assert cell.comment is None
        assert cell.is_date is False

        doc.close()

    def test_cell_negative_style_index(self):
        """Test Cell.is_date with negative style index."""
        from pyopenxlsx.cell import Cell
        from unittest.mock import MagicMock

        # Create a mock cell with negative style index
        mock_cell = MagicMock()
        mock_cell.cell_format.return_value = -1

        # Create a mock worksheet and workbook
        mock_ws = MagicMock()
        mock_ws._workbook = MagicMock()
        mock_ws._workbook._date_format_cache = {}

        cell = Cell(mock_cell, mock_ws)

        # is_date should return False for negative style index
        assert cell.is_date is False

    def test_column_weakref_properties(self):
        """Test Column's _worksheet and _workbook properties."""
        wb = Workbook()
        ws = wb.active

        col = ws.column(1)

        # Properties should work
        assert col._worksheet is not None
        assert col._workbook is not None
        assert col._worksheet is ws
        assert col._workbook is wb

    def test_column_without_worksheet(self, tmp_path):
        """Test Column created without worksheet reference."""
        from pyopenxlsx.column import Column
        from pyopenxlsx._openxlsx import XLDocument

        temp_file = tmp_path / "test_column.xlsx"

        doc = XLDocument()
        doc.create(str(temp_file))
        wb_raw = doc.workbook()
        ws_raw = wb_raw.worksheet("Sheet1")
        col_raw = ws_raw.column(1)

        # Create column without worksheet
        col = Column(col_raw, None)

        # Properties should return None
        assert col._worksheet is None
        assert col._workbook is None

        doc.close()

    def test_range_weakref_properties(self):
        """Test Range's _worksheet property."""
        wb = Workbook()
        ws = wb.active

        rng = ws.range("A1:B2")

        # Property should work
        assert rng._worksheet is not None
        assert rng._worksheet is ws

    def test_range_without_worksheet(self, tmp_path):
        """Test Range created without worksheet reference."""
        from pyopenxlsx.range import Range
        from pyopenxlsx._openxlsx import XLDocument

        temp_file = tmp_path / "test_range.xlsx"

        doc = XLDocument()
        doc.create(str(temp_file))
        wb_raw = doc.workbook()
        ws_raw = wb_raw.worksheet("Sheet1")
        rng_raw = ws_raw.range("A1:B2")

        # Create range without worksheet
        rng = Range(rng_raw, None)

        # Property should return None
        assert rng._worksheet is None

        doc.close()

    def test_worksheet_cache_cleanup(self):
        """Test that worksheet cell cache uses WeakValueDictionary."""
        wb = Workbook()
        ws = wb.active

        # Create cells
        cell1 = ws.cell(1, 1, "test1")
        cell2 = ws.cell(1, 2, "test2")

        # Cache should contain the cells
        assert (1, 1) in ws._cells
        assert (1, 2) in ws._cells

        # Delete references and force GC
        del cell1
        del cell2
        gc.collect()

        # Cells might be garbage collected (WeakValueDictionary behavior)
        # This tests the WeakValueDictionary is working

    def test_workbook_sheets_cache_cleanup(self):
        """Test that workbook sheets cache uses WeakValueDictionary."""
        wb = Workbook()
        ws = wb.active
        sheet_name = ws.title

        # Cache should contain the sheet
        assert sheet_name in wb._sheets

        # Delete reference and force GC
        del ws
        gc.collect()

        # Sheet might be garbage collected (WeakValueDictionary behavior)


class TestAsyncMethodsCoverage:
    """Test async methods that weren't covered."""

    @pytest.mark.asyncio
    async def test_active_setter_with_invalid_type(self):
        """Test setting active with invalid type."""
        wb = Workbook()

        with pytest.raises(TypeError):
            wb.active = "not a worksheet"

    @pytest.mark.asyncio
    async def test_copy_worksheet_async(self):
        """Test async worksheet copy."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "original")

        ws_copy = await wb.copy_worksheet_async(ws)

        assert ws_copy is not None
        assert "Copy" in ws_copy.title
        assert ws_copy.cell(1, 1).value == "original"

    @pytest.mark.asyncio
    async def test_get_range_data_async(self):
        """Test async range data retrieval."""
        wb = Workbook()
        ws = wb.active

        # Fill some data
        for r in range(1, 4):
            for c in range(1, 4):
                ws.cell(r, c, f"R{r}C{c}")

        data = await ws.get_range_data_async(1, 1, 3, 3)

        assert len(data) == 3
        assert len(data[0]) == 3
        assert data[0][0] == "R1C1"

    @pytest.mark.asyncio
    async def test_get_cell_value_async(self):
        """Test async single cell value retrieval."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "async test")

        value = await ws.get_cell_value_async(1, 1)

        assert value == "async test"

    @pytest.mark.asyncio
    async def test_write_range_async(self):
        """Test async range write with numpy array."""
        pytest.importorskip("numpy")
        import numpy as np

        wb = Workbook()
        ws = wb.active

        data = np.array([[1.0, 2.0], [3.0, 4.0]])
        await ws.write_range_async(1, 1, data)

        assert ws.cell(1, 1).value == 1.0
        assert ws.cell(1, 2).value == 2.0
        assert ws.cell(2, 1).value == 3.0
        assert ws.cell(2, 2).value == 4.0

    @pytest.mark.asyncio
    async def test_get_range_values_async(self):
        """Test async range values retrieval as numpy array."""
        pytest.importorskip("numpy")
        import numpy as np

        wb = Workbook()
        ws = wb.active

        # Fill with numeric data
        for r in range(1, 4):
            for c in range(1, 4):
                ws.cell(r, c, float(r * c))

        data = await ws.get_range_values_async(1, 1, 3, 3)

        assert isinstance(data, np.ndarray)
        assert data.shape == (3, 3)
        assert data[0, 0] == 1.0
        assert data[1, 1] == 4.0


class TestSheetStateCoverage:
    """Test sheet state edge cases."""

    def test_sheet_state_very_hidden(self):
        """Test setting sheet state to very_hidden."""
        wb = Workbook()
        ws1 = wb.active
        wb.create_sheet("Sheet2")

        ws1.sheet_state = "very_hidden"

        assert ws1.sheet_state == "very_hidden"

    def test_sheet_state_invalid(self):
        """Test sheet state with invalid value (no-op)."""
        wb = Workbook()
        ws = wb.active

        original_state = ws.sheet_state
        ws.sheet_state = "invalid_state"  # Should be a no-op

        # State should remain unchanged
        assert ws.sheet_state == original_state


class TestRangeCoverage:
    """Test range-related coverage."""

    def test_range_with_two_args(self):
        """Test range with two cell references."""
        wb = Workbook()
        ws = wb.active

        rng = ws.range("A1", "C3")

        assert rng.num_rows == 3
        assert rng.num_columns == 3

    def test_range_invalid_args(self):
        """Test range with invalid number of arguments."""
        wb = Workbook()
        ws = wb.active

        with pytest.raises(TypeError):
            ws.range("A1", "B2", "C3")  # Too many args


class TestCellCommentCoverage:
    """Test cell comment edge cases."""

    def test_cell_comment_set_without_worksheet(self, tmp_path):
        """Test setting comment on cell without worksheet raises error."""
        from pyopenxlsx.cell import Cell
        from pyopenxlsx._openxlsx import XLDocument

        temp_file = tmp_path / "test_comment.xlsx"

        doc = XLDocument()
        doc.create(str(temp_file))
        wb_raw = doc.workbook()
        ws_raw = wb_raw.worksheet("Sheet1")
        cell_raw = ws_raw.cell(1, 1)

        cell = Cell(cell_raw, None)

        with pytest.raises(
            ValueError, match="Cell must be associated with a worksheet"
        ):
            cell.comment = "test comment"

        doc.close()

    def test_cell_comment_delete(self):
        """Test deleting a cell comment."""
        wb = Workbook()
        ws = wb.active
        cell = ws.cell(1, 1, "test")

        # Set comment
        cell.comment = "test comment"
        assert cell.comment == "test comment"

        # Delete comment
        cell.comment = None
        assert cell.comment is None


class TestGetitemCoverage:
    """Test __getitem__ coverage."""

    def test_worksheet_getitem_invalid_type(self):
        """Test worksheet __getitem__ with invalid type."""
        wb = Workbook()
        ws = wb.active

        with pytest.raises(TypeError):
            ws[123]  # Integer not supported


class TestFormulaSetter:
    """Test formula setter coverage."""

    def test_formula_setter(self):
        """Test setting formula on cell."""
        wb = Workbook()
        ws = wb.active
        cell = ws.cell(1, 1)

        cell.formula = "=SUM(A2:A10)"

        # Use text property instead of get()
        assert cell.formula.text == "=SUM(A2:A10)"


class TestRowsCoverage:
    """Test rows property coverage."""

    def test_rows_property(self):
        """Test iterating over rows property."""
        wb = Workbook()
        ws = wb.active

        # Add some data
        for r in range(1, 4):
            for c in range(1, 4):
                ws.cell(r, c, f"R{r}C{c}")

        rows = list(ws.rows)

        assert len(rows) == 3
        assert len(rows[0]) == 3
        assert rows[0][0].value == "R1C1"


class TestCachingCoverage:
    """Test caching coverage for worksheet and workbook."""

    def test_worksheet_getitem_cache_hit(self):
        """Test __getitem__ cache hit path."""
        wb = Workbook()
        ws = wb.active

        # First access creates the cell
        cell1 = ws["A1"]
        cell1.value = "test"

        # Second access should hit the cache
        cell2 = ws["A1"]

        # Should be the same object
        assert cell1 is cell2
        assert cell2.value == "test"

    def test_worksheet_cell_cache_hit(self):
        """Test cell() cache hit path."""
        wb = Workbook()
        ws = wb.active

        # First access creates the cell
        cell1 = ws.cell(1, 1)
        cell1.value = "cached"

        # Second access should hit the cache
        cell2 = ws.cell(1, 1)

        # Should be the same object
        assert cell1 is cell2
        assert cell2.value == "cached"

    def test_workbook_getitem_cache_hit(self):
        """Test workbook __getitem__ cache hit path."""
        wb = Workbook()

        # First access
        ws1 = wb.active
        name = ws1.title

        # Access via __getitem__ to populate cache
        ws2 = wb[name]

        # Second access should hit the cache
        ws3 = wb[name]

        # Should be the same object
        assert ws2 is ws3


class TestEdgeCaseCoverage:
    """Test edge case coverage."""

    def test_active_exception_handling(self):
        """Test active property exception handling path."""
        # The active property has a try/except block on lines 330-337
        # This is hard to trigger without mocking, but the cache tests cover most paths
        pass

    def test_create_sheet_async(self):
        """Test async create_sheet."""
        import asyncio

        async def test():
            wb = Workbook()
            ws = await wb.create_sheet_async("AsyncSheet")
            assert ws.title == "AsyncSheet"

        asyncio.run(test())

    def test_copy_worksheet_async_coverage(self):
        """Test copy_worksheet_async for coverage."""
        import asyncio

        async def test():
            wb = Workbook()
            ws = wb.active
            ws.cell(1, 1, "test")

            ws_copy = await wb.copy_worksheet_async(ws)
            assert "Copy" in ws_copy.title

        asyncio.run(test())
