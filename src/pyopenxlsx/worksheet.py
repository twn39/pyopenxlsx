import asyncio
from weakref import WeakValueDictionary

from ._openxlsx import XLSheetState
from .cell import Cell
from .range import Range
from .merge import MergeCells
from .column import Column
from .data_validation import DataValidations
from .table import Table
from .autofilter import AutoFilter
from .page_setup import PageMargins, PrintOptions, PageSetup


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

    @property
    def has_drawing(self):
        """Check if the worksheet has a drawing (images, charts, etc.)."""
        return self._sheet.has_drawing()

    @property
    def drawing(self):
        """Get the drawing object for the worksheet."""
        return self._sheet.drawing()

    def append(self, iterable):
        row = self.max_row + 1
        values = list(iterable)
        if values:
            self._sheet.write_row_data(row, 1, values)

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

    def insert_row(self, row_number, count=1):
        """Insert one or more rows at the given row number (1-based)."""
        return self._sheet.insert_row(row_number, count)

    def delete_row(self, row_number, count=1):
        """Delete one or more rows starting at the given row number (1-based)."""
        if count == 1:
            return self._sheet.delete_row(row_number)
        return self._sheet.delete_row(row_number, count)

    def insert_column(self, col_number, count=1):
        """Insert one or more columns at the given column number (1-based)."""
        return self._sheet.insert_column(col_number, count)

    def delete_column(self, col_number, count=1):
        """Delete one or more columns starting at the given column number (1-based)."""
        return self._sheet.delete_column(col_number, count)

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
        insert_hyperlinks=False,
        delete_columns=False,
        delete_rows=False,
        select_locked_cells=True,
        select_unlocked_cells=True,
        auto_filter=False,
        sort=False,
        pivot_tables=False,
        format_cells=False,
        format_columns=False,
        format_rows=False,
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
        self._sheet.set_insert_hyperlinks_allowed(insert_hyperlinks)
        self._sheet.set_delete_columns_allowed(delete_columns)
        self._sheet.set_delete_rows_allowed(delete_rows)
        self._sheet.set_select_locked_cells_allowed(select_locked_cells)
        self._sheet.set_select_unlocked_cells_allowed(select_unlocked_cells)
        self._sheet.set_auto_filter_allowed(auto_filter)
        self._sheet.set_sort_allowed(sort)
        self._sheet.set_pivot_tables_allowed(pivot_tables)
        self._sheet.set_format_cells_allowed(format_cells)
        self._sheet.set_format_columns_allowed(format_columns)
        self._sheet.set_format_rows_allowed(format_rows)

    async def protect_async(
        self,
        password=None,
        objects=True,
        scenarios=True,
        insert_columns=False,
        insert_rows=False,
        insert_hyperlinks=False,
        delete_columns=False,
        delete_rows=False,
        select_locked_cells=True,
        select_unlocked_cells=True,
        auto_filter=False,
        sort=False,
        pivot_tables=False,
        format_cells=False,
        format_columns=False,
        format_rows=False,
    ):
        await asyncio.to_thread(
            self.protect,
            password,
            objects,
            scenarios,
            insert_columns,
            insert_rows,
            insert_hyperlinks,
            delete_columns,
            delete_rows,
            select_locked_cells,
            select_unlocked_cells,
            auto_filter,
            sort,
            pivot_tables,
            format_cells,
            format_columns,
            format_rows,
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
            "insert_hyperlinks": self._sheet.insert_hyperlinks_allowed(),
            "delete_columns": self._sheet.delete_columns_allowed(),
            "delete_rows": self._sheet.delete_rows_allowed(),
            "select_locked_cells": self._sheet.select_locked_cells_allowed(),
            "select_unlocked_cells": self._sheet.select_unlocked_cells_allowed(),
            "auto_filter": self._sheet.auto_filter_allowed(),
            "sort": self._sheet.sort_allowed(),
            "pivot_tables": self._sheet.pivot_tables_allowed(),
            "format_cells": self._sheet.format_cells_allowed(),
            "format_columns": self._sheet.format_columns_allowed(),
            "format_rows": self._sheet.format_rows_allowed(),
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

    def add_hyperlink(self, cell_ref, url, tooltip=""):
        """
        Add an external hyperlink to a cell.

        :param cell_ref: Cell reference (e.g., 'A1').
        :param url: URL of the hyperlink.
        :param tooltip: Optional tooltip text.
        """
        self._sheet.add_hyperlink(cell_ref, url, tooltip)

    def add_internal_hyperlink(self, cell_ref, location, tooltip=""):
        """
        Add an internal hyperlink (to another sheet or range) to a cell.

        :param cell_ref: Cell reference (e.g., 'A1').
        :param location: Destination in the workbook (e.g., 'Sheet2!A1').
        :param tooltip: Optional tooltip text.
        """
        self._sheet.add_internal_hyperlink(cell_ref, location, tooltip)

    def has_hyperlink(self, cell_ref):
        """Check if a cell has a hyperlink."""
        return self._sheet.has_hyperlink(cell_ref)

    def get_hyperlink(self, cell_ref):
        """Get the hyperlink target for a cell."""
        return self._sheet.get_hyperlink(cell_ref)

    def remove_hyperlink(self, cell_ref):
        """Remove a hyperlink from a cell."""
        self._sheet.remove_hyperlink(cell_ref)

    def freeze_panes(self, row_or_ref, col=None):
        """
        Freeze the worksheet panes.

        :param row_or_ref: Row number (1-indexed) or a cell reference string (e.g., 'B2').
        :param col: Column number (1-indexed). Only used if row_or_ref is an int.
        """
        if isinstance(row_or_ref, str):
            self._sheet.freeze_panes(row_or_ref)
        elif isinstance(row_or_ref, int):
            if col is None:
                self._sheet.freeze_panes(0, row_or_ref)
            else:
                self._sheet.freeze_panes(col, row_or_ref)
        else:
            raise TypeError("row_or_ref must be an int or a string reference")

    def split_panes(
        self, x_split, y_split, top_left_cell="", active_pane="bottomRight"
    ):
        """
        Split the worksheet panes at given pixel coordinates.

        :param x_split: Horizontal split position in 1/20th of a point.
        :param y_split: Vertical split position in 1/20th of a point.
        :param top_left_cell: Cell address of the top-left cell in the bottom-right pane.
        :param active_pane: The pane that is active ('bottomRight', 'topRight', 'bottomLeft', 'topLeft').
        """
        from ._openxlsx import XLPane

        pane_map = {
            "bottomRight": XLPane.BottomRight,
            "topRight": XLPane.TopRight,
            "bottomLeft": XLPane.BottomLeft,
            "topLeft": XLPane.TopLeft,
        }
        active_pane_enum = pane_map.get(active_pane, XLPane.BottomRight)
        self._sheet.split_panes(x_split, y_split, top_left_cell, active_pane_enum)

    def clear_panes(self):
        """Clear all panes (frozen or split) from the worksheet."""
        self._sheet.clear_panes()

    @property
    def has_panes(self):
        """Check if the worksheet has frozen or split panes."""
        return self._sheet.has_panes()

    @property
    def auto_filter(self):
        """
        Get the AutoFilter object for the worksheet to manage filters.
        Returns None if no AutoFilter is set.
        """
        af = AutoFilter(self._sheet.autofilter_object(), self)
        if not af:
            return None
        return af

    @auto_filter.setter
    def auto_filter(self, value):
        if value is None:
            self._sheet.clear_auto_filter()
        elif isinstance(value, str):
            self._sheet.set_auto_filter(value)
        elif isinstance(value, AutoFilter):
            # If setting an AutoFilter object, just set its reference if it differs
            pass

    @property
    def zoom(self):
        """Get or set the worksheet zoom scale (percentage, e.g., 100)."""
        return self._sheet.zoom()

    @zoom.setter
    def zoom(self, value):
        self._sheet.set_zoom(int(value))

    @property
    def data_validations(self):
        """
        Get the DataValidations object for this worksheet to manage data validation rules.
        """
        return DataValidations(self._sheet.data_validations(), self)

    @property
    def tables(self):
        """
        Get the collection of tables in this worksheet.
        """
        return self._sheet.tables()

    @property
    def table(self):
        """
        Get the first Table object for this worksheet.
        If no table exists, one is created automatically with default name 'Table1' and range 'A1:A1'.
        Note: OpenXLSX now supports multiple tables per worksheet.
        Use the 'tables' property to access all tables or 'add_table' to create new ones.
        """
        tables = self._sheet.tables()
        if len(tables) == 0:
            # Create a default table for backward compatibility
            return self.add_table("Table1", "A1:A1")
        return Table(tables[0], self)

    def add_table(self, name, range_string):
        """
        Add a new table to the worksheet.

        :param name: Table name (no spaces).
        :param range_string: Range reference (e.g., 'A1:C10').
        :return: Table object.
        """
        tables = self._sheet.tables()
        raw_table = tables.add(name, range_string)
        return Table(raw_table, self)

    @property
    def page_margins(self):
        """
        Get the PageMargins object for this worksheet.
        """
        return PageMargins(self._sheet.page_margins(), self)

    @property
    def print_options(self):
        """
        Get the PrintOptions object for this worksheet.
        """
        return PrintOptions(self._sheet.print_options(), self)

    @property
    def page_setup(self):
        """
        Get the PageSetup object for this worksheet.
        """
        return PageSetup(self._sheet.page_setup(), self)

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

        Example::

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

    def stream_writer(self):
        """Get a stream writer for this worksheet."""
        return self._sheet.stream_writer()

    def stream_reader(self):
        """Get a stream reader for this worksheet."""
        return self._sheet.stream_reader()

    def auto_fit_column(self, column_number: int):
        """Auto-fit the specified column."""
        self._sheet.auto_fit_column(column_number)

    def apply_auto_filter(self):
        """Apply auto filter to the worksheet."""
        self._sheet.apply_auto_filter()

    def add_conditional_formatting(self, sqref: str, rule):
        """Add conditional formatting to a range."""
        self._sheet.add_conditional_formatting(sqref, rule)

    def remove_conditional_formatting(self, sqref: str):
        """Remove conditional formatting from a range."""
        self._sheet.remove_conditional_formatting(sqref)

    def clear_all_conditional_formatting(self):
        """Clear all conditional formatting."""
        self._sheet.clear_all_conditional_formatting()

    def set_print_area(self, sqref: str):
        """Set the print area for the worksheet."""
        self._sheet.set_print_area(sqref)

    def set_print_title_rows(self, first_row: int, last_row: int):
        """Set the rows to repeat at top on printed pages."""
        self._sheet.set_print_title_rows(first_row, last_row)

    def set_print_title_cols(self, first_col: int, last_col: int):
        """Set the columns to repeat at left on printed pages."""
        self._sheet.set_print_title_cols(first_col, last_col)

    def add_sparkline(self, location: str, data_range: str, sparkline_type=None):
        """Add a sparkline to the worksheet."""
        if sparkline_type is None:
            self._sheet.add_sparkline(location, data_range)
        else:
            self._sheet.add_sparkline(location, data_range, sparkline_type)

    def add_comment(self, cell_ref: str, text: str, author: str = ""):
        """Add a comment to a cell."""
        self._sheet.add_comment(cell_ref, text, author)
