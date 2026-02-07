"""
Tests for bulk read APIs (get_rows_data, get_row_values, iter_row_values).

These tests verify the optimized C++ bulk read functionality exposed
through pybind11 bindings.
"""

import asyncio
import pytest
from pyopenxlsx import Workbook, load_workbook


class TestBulkReadBasic:
    """Basic tests for bulk read functionality."""

    def test_get_rows_data_empty_sheet(self):
        """Test get_rows_data on an empty worksheet."""
        wb = Workbook()
        ws = wb.active
        result = ws.get_rows_data()
        assert result == []
        wb.close()

    def test_get_rows_data_single_cell(self):
        """Test get_rows_data with a single cell."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "Hello")
        result = ws.get_rows_data()
        assert result == [["Hello"]]
        wb.close()

    def test_get_rows_data_single_row(self):
        """Test get_rows_data with a single row of data."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "A")
        ws.cell(1, 2, "B")
        ws.cell(1, 3, "C")
        result = ws.get_rows_data()
        assert result == [["A", "B", "C"]]
        wb.close()

    def test_get_rows_data_multiple_rows(self):
        """Test get_rows_data with multiple rows."""
        wb = Workbook()
        ws = wb.active
        test_data = [
            ["Name", "Age", "City"],
            ["Alice", 30, "NYC"],
            ["Bob", 25, "LA"],
        ]
        for r, row_data in enumerate(test_data, 1):
            for c, val in enumerate(row_data, 1):
                ws.cell(r, c, val)

        result = ws.get_rows_data()
        assert len(result) == 3
        assert result[0] == ["Name", "Age", "City"]
        assert result[1] == ["Alice", 30, "NYC"]
        assert result[2] == ["Bob", 25, "LA"]
        wb.close()


class TestBulkReadDataTypes:
    """Tests for different data types in bulk read."""

    def test_get_rows_data_string_values(self):
        """Test string values are correctly read."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "Hello World")
        ws.cell(1, 2, "测试中文")
        ws.cell(1, 3, "Special: @#$%^&*()")

        result = ws.get_rows_data()
        assert result[0][0] == "Hello World"
        assert result[0][1] == "测试中文"
        assert result[0][2] == "Special: @#$%^&*()"

    def test_get_rows_data_integer_values(self):
        """Test integer values are correctly read."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, 0)
        ws.cell(1, 2, 42)
        ws.cell(1, 3, -100)
        ws.cell(1, 4, 2147483647)  # Max int32

        result = ws.get_rows_data()
        assert result[0][0] == 0
        assert result[0][1] == 42
        assert result[0][2] == -100
        assert result[0][3] == 2147483647

    def test_get_rows_data_float_values(self):
        """Test float values are correctly read."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, 3.14159)
        ws.cell(1, 2, -2.5)
        ws.cell(1, 3, 0.0)
        ws.cell(1, 4, 1e10)

        result = ws.get_rows_data()
        assert abs(result[0][0] - 3.14159) < 1e-5
        assert result[0][1] == -2.5
        assert result[0][2] == 0.0
        assert result[0][3] == 1e10

    def test_get_rows_data_boolean_values(self):
        """Test boolean values are correctly read."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, True)
        ws.cell(1, 2, False)

        result = ws.get_rows_data()
        assert result[0][0] is True
        assert result[0][1] is False

    def test_get_rows_data_empty_cells(self):
        """Test empty cells return None."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "Start")
        # Leave cell (1, 2) empty
        ws.cell(1, 3, "End")

        result = ws.get_rows_data()
        assert result[0][0] == "Start"
        assert result[0][1] is None
        assert result[0][2] == "End"

    def test_get_rows_data_mixed_types(self):
        """Test mixed data types in the same row."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "Text")
        ws.cell(1, 2, 123)
        ws.cell(1, 3, 45.67)
        ws.cell(1, 4, True)
        # Cell 5 is empty
        ws.cell(1, 6, "End")

        result = ws.get_rows_data()
        assert result[0][0] == "Text"
        assert result[0][1] == 123
        assert abs(result[0][2] - 45.67) < 1e-5
        assert result[0][3] is True
        assert result[0][4] is None
        assert result[0][5] == "End"


class TestGetRowValues:
    """Tests for get_row_values() single row API."""

    def test_get_row_values_first_row(self):
        """Test getting values from the first row."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "A1")
        ws.cell(1, 2, "B1")
        ws.cell(2, 1, "A2")

        result = ws.get_row_values(1)
        assert result == ["A1", "B1"]

    def test_get_row_values_middle_row(self):
        """Test getting values from a middle row."""
        wb = Workbook()
        ws = wb.active
        for r in range(1, 4):
            for c in range(1, 4):
                ws.cell(r, c, f"R{r}C{c}")

        result = ws.get_row_values(2)
        assert result == ["R2C1", "R2C2", "R2C3"]

    def test_get_row_values_last_row(self):
        """Test getting values from the last row."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "First")
        ws.cell(5, 1, "Last")
        ws.cell(5, 2, 100)

        result = ws.get_row_values(5)
        assert result == ["Last", 100]

    def test_get_row_values_with_gaps(self):
        """Test row values with empty cells in between."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "A")
        ws.cell(1, 3, "C")
        ws.cell(1, 5, "E")

        result = ws.get_row_values(1)
        assert len(result) == 5
        assert result[0] == "A"
        assert result[1] is None
        assert result[2] == "C"
        assert result[3] is None
        assert result[4] == "E"


class TestIterRowValues:
    """Tests for iter_row_values() iterator API."""

    def test_iter_row_values_empty_sheet(self):
        """Test iteration on empty sheet yields nothing."""
        wb = Workbook()
        ws = wb.active
        result = list(ws.iter_row_values())
        assert result == []

    def test_iter_row_values_single_row(self):
        """Test iteration with single row."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "Only")

        rows = list(ws.iter_row_values())
        assert len(rows) == 1
        assert rows[0] == ["Only"]

    def test_iter_row_values_multiple_rows(self):
        """Test iteration over multiple rows."""
        wb = Workbook()
        ws = wb.active
        test_data = [
            ["Header1", "Header2"],
            [1, 2],
            [3, 4],
            [5, 6],
        ]
        for r, row_data in enumerate(test_data, 1):
            for c, val in enumerate(row_data, 1):
                ws.cell(r, c, val)

        rows = list(ws.iter_row_values())
        assert len(rows) == 4
        assert rows[0] == ["Header1", "Header2"]
        assert rows[1] == [1, 2]
        assert rows[2] == [3, 4]
        assert rows[3] == [5, 6]

    def test_iter_row_values_is_generator(self):
        """Test that iter_row_values returns a generator."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "Test")

        iterator = ws.iter_row_values()
        # Check it's a generator
        assert hasattr(iterator, "__iter__")
        assert hasattr(iterator, "__next__")


class TestBulkReadConsistency:
    """Tests to verify consistency between bulk read and cell-by-cell read."""

    def test_get_rows_data_matches_cell_read(self):
        """Verify get_rows_data matches cell-by-cell reading."""
        wb = Workbook()
        ws = wb.active

        # Create test data
        test_data = [
            ["Name", "Age", "Score"],
            ["Alice", 30, 95.5],
            ["Bob", 25, 88.0],
            ["Charlie", 35, 92.3],
        ]
        for r, row_data in enumerate(test_data, 1):
            for c, val in enumerate(row_data, 1):
                ws.cell(r, c, val)

        # Get bulk data
        bulk_result = ws.get_rows_data()

        # Compare with cell-by-cell
        for r in range(len(test_data)):
            for c in range(len(test_data[r])):
                cell_value = ws.cell(r + 1, c + 1).value
                bulk_value = bulk_result[r][c]

                if isinstance(cell_value, float):
                    assert abs(cell_value - bulk_value) < 1e-5
                else:
                    assert cell_value == bulk_value

    def test_get_row_values_matches_get_rows_data(self):
        """Verify get_row_values matches corresponding row in get_rows_data."""
        wb = Workbook()
        ws = wb.active

        for r in range(1, 5):
            for c in range(1, 4):
                ws.cell(r, c, r * 10 + c)

        all_rows = ws.get_rows_data()

        for r in range(1, 5):
            single_row = ws.get_row_values(r)
            assert single_row == all_rows[r - 1]

    def test_iter_row_values_matches_get_rows_data(self):
        """Verify iter_row_values yields same data as get_rows_data."""
        wb = Workbook()
        ws = wb.active

        for r in range(1, 4):
            for c in range(1, 3):
                ws.cell(r, c, f"({r},{c})")

        all_rows = ws.get_rows_data()
        iter_rows = list(ws.iter_row_values())

        assert all_rows == iter_rows


class TestBulkReadFilePersistence:
    """Tests for bulk read with file save/load cycle."""

    def test_get_rows_data_after_save_load(self, tmp_path):
        """Test bulk read after saving and loading a file."""
        file_path = tmp_path / "test_bulk.xlsx"

        # Create and save
        wb = Workbook()
        ws = wb.active
        test_data = [
            ["ID", "Name", "Value"],
            [1, "Test", 100.5],
            [2, "Data", 200.75],
        ]
        for r, row_data in enumerate(test_data, 1):
            for c, val in enumerate(row_data, 1):
                ws.cell(r, c, val)

        wb.save(str(file_path))
        wb.close()

        # Load and verify
        wb2 = load_workbook(str(file_path))
        ws2 = wb2.active
        result = ws2.get_rows_data()

        assert len(result) == 3
        assert result[0] == ["ID", "Name", "Value"]
        assert result[1][0] == 1
        assert result[1][1] == "Test"
        assert abs(result[1][2] - 100.5) < 1e-5

        wb2.close()


class TestAsyncBulkRead:
    """Tests for async versions of bulk read APIs."""

    @pytest.mark.asyncio
    async def test_get_rows_data_async(self):
        """Test async get_rows_data."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "Async")
        ws.cell(1, 2, "Test")
        ws.cell(2, 1, 100)
        ws.cell(2, 2, 200)

        result = await ws.get_rows_data_async()

        assert len(result) == 2
        assert result[0] == ["Async", "Test"]
        assert result[1] == [100, 200]

        await wb.close_async()

    @pytest.mark.asyncio
    async def test_get_row_values_async(self):
        """Test async get_row_values."""
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, "Row1")
        ws.cell(1, 2, 10)
        ws.cell(2, 1, "Row2")
        ws.cell(2, 2, 42)

        row1 = await ws.get_row_values_async(1)
        row2 = await ws.get_row_values_async(2)

        assert row1 == ["Row1", 10]
        assert row2 == ["Row2", 42]

        await wb.close_async()

    @pytest.mark.asyncio
    async def test_async_matches_sync(self):
        """Verify async results match sync results."""
        wb = Workbook()
        ws = wb.active

        for r in range(1, 4):
            for c in range(1, 4):
                ws.cell(r, c, r * c)

        sync_result = ws.get_rows_data()
        async_result = await ws.get_rows_data_async()

        assert sync_result == async_result

        await wb.close_async()

    @pytest.mark.asyncio
    async def test_async_concurrent_reads(self):
        """Test concurrent async reads."""
        wb = Workbook()
        ws = wb.active

        for r in range(1, 6):
            ws.cell(r, 1, f"Row{r}")
            ws.cell(r, 2, r * 10)

        # Run multiple async reads concurrently
        results = await asyncio.gather(
            ws.get_row_values_async(1),
            ws.get_row_values_async(2),
            ws.get_row_values_async(3),
            ws.get_row_values_async(4),
            ws.get_row_values_async(5),
        )

        assert len(results) == 5
        assert results[0] == ["Row1", 10]
        assert results[2] == ["Row3", 30]
        assert results[4] == ["Row5", 50]

        await wb.close_async()


class TestBulkReadLargeData:
    """Performance-related tests for bulk read with larger datasets."""

    def test_get_rows_data_100_rows(self):
        """Test bulk read with 100 rows."""
        wb = Workbook()
        ws = wb.active

        num_rows = 100
        num_cols = 10

        for r in range(1, num_rows + 1):
            for c in range(1, num_cols + 1):
                ws.cell(r, c, r * 100 + c)

        result = ws.get_rows_data()

        assert len(result) == num_rows
        assert len(result[0]) == num_cols
        assert result[0][0] == 101  # 1*100 + 1
        assert result[99][9] == 10010  # 100*100 + 10

    def test_get_rows_data_wide_sheet(self):
        """Test bulk read with many columns."""
        wb = Workbook()
        ws = wb.active

        num_cols = 50
        for c in range(1, num_cols + 1):
            ws.cell(1, c, f"Col{c}")
            ws.cell(2, c, c)

        result = ws.get_rows_data()

        assert len(result) == 2
        assert len(result[0]) == num_cols
        assert result[0][0] == "Col1"
        assert result[0][49] == "Col50"
        assert result[1][0] == 1
        assert result[1][49] == 50
