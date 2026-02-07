import asyncio
from weakref import WeakValueDictionary

from ._openxlsx import XLSheetState
from .cell import Cell
from .range import Range
from .merge import MergeCells
from .column import Column


class Worksheet:
    """
    Represents an Excel worksheet.

    Uses WeakValueDictionary for cell caching to allow garbage collection
    of Cell objects when they are no longer referenced elsewhere.
    """

    def __init__(self, raw_sheet, workbook=None):
        self._sheet = raw_sheet
        self._workbook = workbook
        # Use WeakValueDictionary to avoid keeping Cell objects alive indefinitely
        # Cells will be garbage collected when no external references remain
        self._cells = WeakValueDictionary()

    @property
    def title(self):
        return self._sheet.name()

    @title.setter
    def title(self, value):
        self._sheet.set_name(value)

    @property
    def index(self):
        return self._sheet.index() - 1

    @index.setter
    def index(self, value):
        self._sheet.set_index(value + 1)

    @property
    def sheet_state(self):
        state = self._sheet.visibility()
        if state == XLSheetState.Visible:
            return "visible"
        elif state == XLSheetState.Hidden:
            return "hidden"
        elif state == XLSheetState.VeryHidden:
            return "very_hidden"
        return "visible"

    @sheet_state.setter
    def sheet_state(self, value):
        if value == "visible":
            self._sheet.set_visibility(XLSheetState.Visible)
        elif value == "hidden":
            self._sheet.set_visibility(XLSheetState.Hidden)
        elif value == "very_hidden":
            self._sheet.set_visibility(XLSheetState.VeryHidden)

    @property
    def max_row(self):
        return self._sheet.row_count()

    @property
    def max_column(self):
        return self._sheet.column_count()

    def append(self, iterable):
        row = self.max_row + 1
        for col, val in enumerate(iterable, start=1):
            self.cell(row, col, value=val)

    async def append_async(self, iterable):
        await asyncio.to_thread(self.append, iterable)

    @property
    def rows(self):
        for r in range(1, self.max_row + 1):
            yield tuple(self.cell(r, c) for c in range(1, self.max_column + 1))

    def __getitem__(self, key):
        if isinstance(key, str):
            if key in self._cells:
                return self._cells[key]
            c = Cell(self._sheet.cell(key), self)
            self._cells[key] = c
            return c
        raise TypeError("Only string references (e.g., 'A1') are supported")

    def cell(self, row, column, value=None):
        key = (row, column)
        if key in self._cells:
            c = self._cells[key]
        else:
            c = Cell(self._sheet.cell(row, column), self)
            self._cells[key] = c

        if value is not None:
            c.value = value
        return c

    def _get_cached_cell(self, raw_cell):
        """Internal helper to get a cached Cell object from a raw XLCell."""
        ref = raw_cell.cell_reference()
        key = (ref.row(), ref.column())
        if key in self._cells:
            return self._cells[key]
        c = Cell(raw_cell, self)
        self._cells[key] = c
        return c

    def range(self, *args):
        if len(args) == 1:
            return Range(self._sheet.range(args[0]), self)
        elif len(args) == 2:
            return Range(self._sheet.range(args[0], args[1]), self)
        raise TypeError("range() takes 1 or 2 arguments")

    def merge_cells(self, range_string):
        self._sheet.merge_cells(range_string)

    async def merge_cells_async(self, range_string):
        await asyncio.to_thread(self.merge_cells, range_string)

    def unmerge_cells(self, range_string):
        self._sheet.unmerge_cells(range_string)

    async def unmerge_cells_async(self, range_string):
        await asyncio.to_thread(self.unmerge_cells, range_string)

    def set_column_format(self, column, style_index):
        if isinstance(column, int):
            self._sheet.set_column_format(column, style_index)
        else:
            self._sheet.set_column_format(str(column), style_index)

    def set_row_format(self, row, style_index):
        self._sheet.set_row_format(row, style_index)

    @property
    def merges(self):
        return MergeCells(self._sheet.merges())

    def column(self, col):
        """
        Get a Column object.
        """
        if isinstance(col, int):
            return Column(self._sheet.column(col), self)
        return Column(self._sheet.column(str(col)), self)

    def protect(
        self,
        password=None,
        objects=True,
        scenarios=True,
        insert_columns=False,
        insert_rows=False,
        delete_columns=False,
        delete_rows=False,
        select_locked_cells=True,
        select_unlocked_cells=True,
    ):
        """
        Protect the worksheet.
        """
        if password:
            self._sheet.set_password(password)
        self._sheet.protect_sheet(True)
        self._sheet.protect_objects(objects)
        self._sheet.protect_scenarios(scenarios)
        self._sheet.set_insert_columns_allowed(insert_columns)
        self._sheet.set_insert_rows_allowed(insert_rows)
        self._sheet.set_delete_columns_allowed(delete_columns)
        self._sheet.set_delete_rows_allowed(delete_rows)
        self._sheet.set_select_locked_cells_allowed(select_locked_cells)
        self._sheet.set_select_unlocked_cells_allowed(select_unlocked_cells)

    async def protect_async(
        self,
        password=None,
        objects=True,
        scenarios=True,
        insert_columns=False,
        insert_rows=False,
        delete_columns=False,
        delete_rows=False,
        select_locked_cells=True,
        select_unlocked_cells=True,
    ):
        await asyncio.to_thread(
            self.protect,
            password,
            objects,
            scenarios,
            insert_columns,
            insert_rows,
            delete_columns,
            delete_rows,
            select_locked_cells,
            select_unlocked_cells,
        )

    def unprotect(self):
        """
        Unprotect the worksheet.
        """
        self._sheet.protect_sheet(False)
        self._sheet.clear_password()

    async def unprotect_async(self):
        await asyncio.to_thread(self.unprotect)

    @property
    def protection(self):
        """
        Get the protection status of the worksheet.
        """
        return {
            "protected": self._sheet.sheet_protected(),
            "password_set": self._sheet.password_is_set(),
            "objects": self._sheet.objects_protected(),
            "scenarios": self._sheet.scenarios_protected(),
            "insert_columns": self._sheet.insert_columns_allowed(),
            "insert_rows": self._sheet.insert_rows_allowed(),
            "delete_columns": self._sheet.delete_columns_allowed(),
            "delete_rows": self._sheet.delete_rows_allowed(),
            "select_locked_cells": self._sheet.select_locked_cells_allowed(),
            "select_unlocked_cells": self._sheet.select_unlocked_cells_allowed(),
        }

    def add_image(self, img_path, anchor="A1", width=None, height=None):
        """
        Add an image to the worksheet.

        :param img_path: Path to the image file.
        :param anchor: Cell reference for the top-left corner of the image (e.g., 'A1').
        :param width: Width of the image in pixels. If None, it will try to get it from the image.
        :param height: Height of the image in pixels. If None, it will try to get it from the image.
        """
        from pathlib import Path

        img_path = Path(img_path)
        if not img_path.exists():
            raise FileNotFoundError(f"Image file not found: {img_path}")

        extension = img_path.suffix.lower().lstrip(".")
        if extension not in ["png", "jpg", "jpeg", "gif"]:
            raise ValueError(f"Unsupported image format: {extension}")

        # Normalize extension for OOXML
        if extension == "jpeg":
            extension = "jpg"

        with open(img_path, "rb") as f:
            img_data = f.read()

        if width is None or height is None:
            try:
                from PIL import Image

                with Image.open(img_path) as img:
                    w, h = img.size
                    if width is None and height is None:
                        width = w
                        height = h
                    elif width is not None and height is None:
                        height = int(h * (width / w))
                    elif width is None and height is not None:
                        width = int(w * (height / h))
            except ImportError:
                if width is None or height is None:
                    raise ImportError(
                        "Pillow is required to automatically detect image dimensions. "
                        "Please install it or provide width and height manually."
                    )

        # Parse anchor
        from ._openxlsx import XLCellReference

        ref = XLCellReference(anchor)

        if width is None or height is None:
            raise ValueError("Width and height must be provided or detected.")

        self._sheet.add_image(
            img_data, extension, ref.row(), ref.column(), int(width), int(height)
        )

    async def add_image_async(self, img_path, anchor="A1", width=None, height=None):
        await asyncio.to_thread(self.add_image, img_path, anchor, width, height)

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
        :param value: Value to set (str, int, float, bool, or None)

        Example:
            # Fast bulk write
            for r in range(1, 1001):
                for c in range(1, 51):
                    ws.set_cell_value(r, c, f"R{r}C{c}")
        """
        self._sheet.set_cell_value(row, column, value)

    async def set_cell_value_async(self, row: int, column: int, value):
        """Async version of set_cell_value()."""
        await asyncio.to_thread(self.set_cell_value, row, column, value)

    def write_rows(self, start_row: int, data, start_col: int = 1):
        """
        Write a 2D Python list to a worksheet range.

        This is optimized for any Python data (strings, mixed types, etc.).
        For pure numeric data, use write_range() with numpy for best performance.

        :param start_row: Starting row number (1-indexed)
        :param data: 2D list/tuple of values [[row1_val1, row1_val2, ...], [row2_val1, ...], ...]
        :param start_col: Starting column number (1-indexed), defaults to 1

        Example:
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

        Example:
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
