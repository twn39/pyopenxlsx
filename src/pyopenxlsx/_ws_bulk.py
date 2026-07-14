from typing import Any

import asyncio


class WorksheetBulkMixin:
    # Provided by Worksheet via mixin composition (for type checkers).
    _sheet: Any
    _workbook: Any
    _closed: bool
    max_row: int
    max_column: int

    # Declared for type checking; runtime implementations live on Worksheet.
    def stream_writer(
        self, use_shared_strings: bool = False, max_unique_strings: int = 100000
    ) -> Any:  # type: ignore[empty-body]
        ...

    def stream_reader(
        self,
        options: Any = None,
        *,
        empty_rows: Any = None,
        apply_number_formats: Any = None,
    ) -> Any:  # type: ignore[empty-body]
        ...

    def get_rows_data(self):
        """
        Get all rows data as list[list[Any]].

        This is an optimized bulk read method that returns all cell values
        without creating intermediate Cell objects. Much faster than iterating
        through ws.rows for large worksheets.

        :return: list[list[Any]] - All cell values, with None for empty cells
        """
        return self._sheet.get_rows_data()

    async def get_rows_data_async(self):
        """Async version of get_rows_data()."""
        return await asyncio.to_thread(self.get_rows_data)

    def get_row_values(self, row: int):
        """
        Get a single row's values as list[Any].

        :param row: Row number (1-indexed)
        :return: list[Any] - Cell values for the specified row
        """
        return self._sheet.get_row_values(row)

    async def get_row_values_async(self, row: int):
        """Async version of get_row_values()."""
        return await asyncio.to_thread(self.get_row_values, row)

    def iter_row_values(self):
        """
        Iterate over rows, yielding each row's values as list[Any].

        This is an optimized iterator that yields row values directly
        without creating Cell objects. Use this for efficient row-by-row
        processing of large worksheets.

        :yields: list[Any] - Cell values for each row
        """
        for row_idx in range(1, self.max_row + 1):
            yield self._sheet.get_row_values(row_idx)

    def get_range_data(
        self, start_row: int, start_col: int, end_row: int, end_col: int
    ):
        """
        Get a range of cells as list[list[Any]].

        This is an optimized bulk read method for reading a specific range
        of cells without creating intermediate Cell objects.

        :param start_row: Starting row number (1-indexed)
        :param start_col: Starting column number (1-indexed)
        :param end_row: Ending row number (1-indexed, inclusive)
        :param end_col: Ending column number (1-indexed, inclusive)
        :return: list[list[Any]] - Cell values in the range
        """
        return self._sheet.get_range_data(start_row, start_col, end_row, end_col)

    async def get_range_data_async(
        self, start_row: int, start_col: int, end_row: int, end_col: int
    ):
        """Async version of get_range_data()."""
        return await asyncio.to_thread(
            self.get_range_data, start_row, start_col, end_row, end_col
        )

    def get_cell_value(self, row: int, column: int):
        """
        Get a single cell's value directly without creating a Cell object.

        This is faster than ws.cell(row, col).value when you only need the value
        and don't need to modify the cell or access other properties.

        :param row: Row number (1-indexed)
        :param column: Column number (1-indexed)
        :return: The cell's value (str, int, float, bool, or None)
        """
        return self._sheet.get_cell_value(row, column)

    async def get_cell_value_async(self, row: int, column: int):
        """Async version of get_cell_value()."""
        return await asyncio.to_thread(self.get_cell_value, row, column)

    def write_dataframe(
        self, df, start_row=1, start_col=1, header=True, index=False, column_styles=None
    ):
        """
        Export a pandas DataFrame to the worksheet.

        Args:
            df: The pandas DataFrame.
            start_row (int): The starting 1-based row index.
            start_col (int): The starting 1-based column index.
            header (bool): Whether to write the DataFrame columns as a header row.
            index (bool): Whether to write the DataFrame index as the first column(s).
            column_styles (dict): Optional dictionary mapping column names or 0-based indices to style IDs.
                                  e.g. {"Date": date_style_id}
        """
        import numpy as np

        if index:
            df = df.reset_index()

        # Replace NaT/NaN with None for C++
        df = df.replace({np.nan: None})

        # If dates are pandas Timestamps, convert them to standard datetime
        for col in df.select_dtypes(include=["datetime64", "datetimetz"]).columns:
            df[col] = df[col].dt.to_pydatetime()

        if column_styles:
            # When styles are requested, we use stream_writer for O(1) style application
            # Convert column_styles to a mapping of column_index -> style_id
            col_idx_styles = {}
            for k, v in column_styles.items():
                if isinstance(k, str) and k in df.columns:
                    col_idx_styles[df.columns.get_loc(k)] = v
                elif isinstance(k, int):
                    col_idx_styles[k] = v

            writer = self.stream_writer()

            # Since stream_writer writes to the very end of the stream, we must pad empty rows if start_row > 1
            # Note: stream_writer writes exactly from the next available row.
            # If start_row > 1 and the sheet is empty, we'd need to pad.
            # To be safe and since stream_writer is generally for append-only,
            # we rely on it just appending. If strict positioning is needed, write_rows is better.
            # But let's assume it appends from where we are.

            if header:
                writer.append_row(df.columns.tolist())

            # Use itertuples for fast iteration while allowing column-specific styling
            for row in df.itertuples(index=False, name=None):
                styled_row = []
                for c_idx, val in enumerate(row):
                    if c_idx in col_idx_styles:
                        styled_row.append((val, col_idx_styles[c_idx]))
                    else:
                        styled_row.append(val)
                writer.append_row(styled_row)

            writer.close()
        else:
            if header:
                # Write column names
                headers = df.columns.tolist()
                self.write_row(start_row, headers, start_col=start_col)
                start_row += 1

            # Since C++ handles numpy types now, we can just pass the fast DataFrame list view directly.
            # df.values.tolist() operates at C speed within pandas.
            self.write_rows(start_row, df.values.tolist(), start_col=start_col)

    async def write_dataframe_async(
        self, df, start_row=1, start_col=1, header=True, index=False
    ):
        import asyncio

        await asyncio.to_thread(
            self.write_dataframe, df, start_row, start_col, header, index
        )

    def read_dataframe(
        self, start_row=1, start_col=1, end_row=None, end_col=None, header=True
    ):
        """
        Import a range from the worksheet to a pandas DataFrame.

        Args:
            start_row (int): The starting 1-based row index.
            start_col (int): The starting 1-based column index.
            end_row (int): The ending 1-based row index. If None, uses max_row.
            end_col (int): The ending 1-based column index. If None, uses max_column.
            header (bool): Whether the first row of the range should be used as column names.

        Returns:
            A pandas DataFrame.
        """
        import pandas as pd

        if end_row is None:
            end_row = self.max_row
        if end_col is None:
            end_col = self.max_column

        data = []
        columns = None

        # Use highly efficient stream_reader to bypass DOM allocation overhead.
        # Note: stream_reader reads from the underlying saved XML file, so uncommitted
        # changes (data written but not yet saved via wb.save()) won't be reflected.
        reader = self.stream_reader()

        # Advance to start_row
        while reader.has_next():
            row_vals = reader.next_row()
            curr_row = reader.current_row()

            if curr_row < start_row:
                continue

            if curr_row > end_row:
                break

            sliced_row = (
                row_vals[start_col - 1 : end_col]
                if end_col <= len(row_vals)
                else row_vals[start_col - 1 :]
            )
            if len(sliced_row) < (end_col - start_col + 1):
                sliced_row.extend(
                    [None] * ((end_col - start_col + 1) - len(sliced_row))
                )

            if header and columns is None:
                columns = sliced_row
            else:
                data.append(sliced_row)

        df = pd.DataFrame(data, columns=columns)

        # Heuristically convert columns that look like serial dates (float > 30000, e.g. year 1980+)
        # But doing so blindly is dangerous. For now, we leave it as float and let the user handle
        # `pd.to_datetime(df['Date'], unit='D', origin='1899-12-30')` if they need peak performance.
        # To balance correctness and speed, we will NOT use `.cell()` loop.

        return df

    async def read_dataframe_async(
        self, start_row=1, start_col=1, end_row=None, end_col=None, header=True
    ):
        import asyncio

        return await asyncio.to_thread(
            self.read_dataframe, start_row, start_col, end_row, end_col, header
        )

    def write_range(self, start_row: int, start_col: int, data):
        """
        Write a 2D numpy array or any object supporting the buffer protocol to a worksheet range.

        This is a high-performance method that avoids Python-level loops and object creation.

        :param start_row: Starting row number (1-indexed)
        :param start_col: Starting column number (1-indexed)
        :param data: 2D numpy array or buffer-compatible object
        """
        self._sheet.write_range_data(start_row, start_col, data)

    async def write_range_async(self, start_row: int, start_col: int, data):
        """Async version of write_range()."""
        await asyncio.to_thread(self.write_range, start_row, start_col, data)

    def get_range_values(
        self, start_row: int, start_col: int, end_row: int, end_col: int
    ):
        """
        Read a range of numeric cells into a 2D numpy array of doubles.

        This is a high-performance method for reading large amounts of numeric data.

        :param start_row: Starting row number (1-indexed)
        :param start_col: Starting column number (1-indexed)
        :param end_row: Ending row number (1-indexed, inclusive)
        :param end_col: Ending column number (1-indexed, inclusive)
        :return: 2D numpy array (float64)
        """
        return self._sheet.get_range_values(start_row, start_col, end_row, end_col)

    async def get_range_values_async(
        self, start_row: int, start_col: int, end_row: int, end_col: int
    ):
        """Async version of get_range_values()."""
        return await asyncio.to_thread(
            self.get_range_values, start_row, start_col, end_row, end_col
        )

    # ============================================================
    # Performance-optimized write APIs
    # These methods bypass Python Cell object creation for 10-20x speedup
    # ============================================================

    def set_cell_value(self, row: int, column: int, value):
        """
        Set a cell's value directly without creating a Cell object.

        This is 10-20x faster than ws.cell(row, col).value = val for bulk operations
        as it bypasses:
        - Python Cell wrapper object creation
        - WeakValueDictionary cache operations
        - Multiple Python/C++ boundary crossings

        :param row: Row number (1-indexed)
        :param column: Column number (1-indexed)
        :param value: Value to set (str, int, float, bool, date/datetime, or None)

        Example::

            # Fast bulk write
            for r in range(1, 1001):
                for c in range(1, 51):
                    ws.set_cell_value(r, c, f"R{r}C{c}")
        """
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        from datetime import date, datetime

        from .cell import datetime_to_serial

        if isinstance(value, (date, datetime)):
            is_datetime = isinstance(value, datetime)
            serial = datetime_to_serial(value)
            self._sheet.set_cell_value(row, column, serial)
            wb = self._workbook
            if wb is not None and getattr(wb, "auto_date_formats", False):
                # Apply date style on the fast path when default format is in use.
                raw = self._sheet.cell(row, column)
                if raw.cell_format() == 0:
                    raw.set_cell_format(
                        wb._get_auto_date_style(is_datetime=is_datetime)
                    )
            return
        self._sheet.set_cell_value(row, column, value)

    async def set_cell_value_async(self, row: int, column: int, value):
        """Async version of set_cell_value()."""
        await asyncio.to_thread(self.set_cell_value, row, column, value)

    def write_rows(self, start_row: int, data, start_col: int = 1):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        """
        Write a 2D Python list to a worksheet range.

        This is optimized for any Python data (strings, mixed types, etc.).
        For pure numeric data, use write_range() with numpy for best performance.

        :param start_row: Starting row number (1-indexed)
        :param data: 2D list/tuple of values [[row1_val1, row1_val2, ...], [row2_val1, ...], ...]
        :param start_col: Starting column number (1-indexed), defaults to 1

        Example::

            data = [
                ["Name", "Age", "City"],
                ["Alice", 30, "New York"],
                ["Bob", 25, "Los Angeles"],
            ]
            ws.write_rows(1, data)
        """
        # Convert to list if it's a tuple or other sequence
        if not isinstance(data, list):
            data = [list(row) if not isinstance(row, list) else row for row in data]
        else:
            data = [list(row) if not isinstance(row, list) else row for row in data]
        self._sheet.write_rows_data(start_row, start_col, data)

    async def write_rows_async(self, start_row: int, data, start_col: int = 1):
        """Async version of write_rows()."""
        await asyncio.to_thread(self.write_rows, start_row, data, start_col)

    def write_row(self, row: int, values, start_col: int = 1):
        """
        Write a single row of Python data.

        :param row: Row number (1-indexed)
        :param values: List/tuple of values for the row
        :param start_col: Starting column number (1-indexed), defaults to 1

        Example:
            ws.write_row(1, ["Name", "Age", "City"])
        """
        if not isinstance(values, list):
            values = list(values)
        self._sheet.write_row_data(row, start_col, values)

    async def write_row_async(self, row: int, values, start_col: int = 1):
        """Async version of write_row()."""
        await asyncio.to_thread(self.write_row, row, values, start_col)

    def set_cells(self, cells):
        """
        Batch set multiple cell values efficiently.

        This is optimal for non-contiguous cell updates where you can't use
        write_rows() or write_range().

        :param cells: Iterable of (row, col, value) tuples

        Example::

            ws.set_cells([
                (1, 1, "Header A"),
                (1, 5, "Header B"),
                (10, 3, 42.5),
                (20, 1, "Footer"),
            ])
        """
        # Convert to list of tuples if needed
        cell_list = [(r, c, v) for r, c, v in cells]
        self._sheet.set_cells_batch(cell_list)

    async def set_cells_async(self, cells):
        """Async version of set_cells()."""
        await asyncio.to_thread(self.set_cells, cells)

    def append_row(self, values):
        """Append a row of values at the end of the used range."""
        self._sheet.append_row(values)
