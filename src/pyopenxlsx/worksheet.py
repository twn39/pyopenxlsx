import asyncio
from collections import deque
from weakref import WeakValueDictionary

from ._openxlsx import XLSheetState
from .cell import Cell
from .range import Range
from .merge import MergeCells
from .column import Column
from ._ws_bulk import WorksheetBulkMixin
from ._ws_drawing import WorksheetDrawingMixin
from ._ws_features import WorksheetFeaturesMixin
from ._ws_page import WorksheetPageMixin
from ._ws_protection import WorksheetProtectionMixin


class Worksheet(
    WorksheetBulkMixin,
    WorksheetDrawingMixin,
    WorksheetFeaturesMixin,
    WorksheetPageMixin,
    WorksheetProtectionMixin,
):
    """
    Represents an Excel worksheet.

    Uses WeakValueDictionary for cell caching to allow garbage collection
    of Cell objects when they are no longer referenced elsewhere.

    Performance note
    ----------------
    Prefer bulk APIs (``set_cell_value``, ``write_rows``, ``set_cells``,
    ``write_range``, ``get_range_values``) for hot loops. Per-cell ``Cell``
    wrappers allocate Python objects and are best for sparse edits.
    """

    def __init__(self, raw_sheet, workbook=None):
        self._sheet = raw_sheet
        self._workbook = workbook
        # Use WeakValueDictionary to avoid keeping Cell objects alive indefinitely
        # Cells will be garbage collected when no external references remain
        self._cells = WeakValueDictionary()
        # Performance optimization: Keep strong references to the last N accessed cells.
        # This prevents immediate garbage collection of transient Cell objects in loops
        # while keeping the overall cache clean and safe.
        self._cells_mru = deque(maxlen=64)

    @property
    def _closed(self):
        """Return True if the parent workbook has been closed."""
        return self._workbook._closed if self._workbook is not None else False

    @property
    def title(self):
        return self._sheet.name()

    @title.setter
    def title(self, value):
        self._sheet.set_name(value)

    @property
    def name(self):
        """Alias for title to maintain compatibility and prevent dynamic attribute bugs."""
        return self.title

    @name.setter
    def name(self, value):
        self.title = value

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

    def iter_rows(
        self,
        min_row=None,
        max_row=None,
        min_col=None,
        max_col=None,
        values_only=False,
    ):
        """Iterate over worksheet rows, openpyxl-compatible.

        Args:
            min_row: First row (1-based, inclusive). Defaults to 1.
            max_row: Last row (1-based, inclusive). Defaults to ws.max_row.
            min_col: First column (1-based, inclusive). Defaults to 1.
            max_col: Last column (1-based, inclusive). Defaults to ws.max_column.
            values_only: If True, yield tuples of raw Python values without
                creating any Cell objects.  Uses the C++ fast-read path
                (one get_row_values() call per row), so peak memory stays
                O(columns) regardless of total row count.
                If False, yield tuples of Cell objects (same semantics as
                ws.rows, but with configurable range bounds).

        Yields:
            tuple[Any, ...] when values_only=True
            tuple[Cell, ...] when values_only=False

        Example::

            # openpyxl-compatible value scan (fast):
            for row in ws.iter_rows(values_only=True):
                process(row)

            # Partial range, Cell objects:
            for row in ws.iter_rows(min_row=2, max_row=100, min_col=1, max_col=5):
                for cell in row:
                    print(cell.value)
        """
        _min_row = min_row if min_row is not None else 1
        _max_row = max_row if max_row is not None else self.max_row
        _min_col = min_col if min_col is not None else 1
        _max_col = max_col if max_col is not None else self.max_column

        if values_only:
            # Fast path: one C++ get_row_values() call per row, zero Cell /
            # WeakRef allocations.  The C++ side returns all columns for the
            # row; we trim to [_min_col, _max_col] and pad with None if the
            # row is shorter than the requested range.
            need = _max_col - _min_col + 1
            for r in range(_min_row, _max_row + 1):
                raw = self._sheet.get_row_values(r)
                sliced = raw[_min_col - 1 : _max_col]
                if len(sliced) < need:
                    sliced = sliced + [None] * (need - len(sliced))
                yield tuple(sliced)
        else:
            # Cell-object path: identical semantics to ws.rows but
            # honours the requested row/column bounds.
            for r in range(_min_row, _max_row + 1):
                yield tuple(self.cell(r, c) for c in range(_min_col, _max_col + 1))

    @property
    def rows(self):
        """Iterate over all rows as tuples of Cell objects.

        For value-only iteration (10-20x faster for read-only access), use::

            ws.iter_rows(values_only=True)
        """
        return self.iter_rows()

    def __getitem__(self, key):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        if isinstance(key, str):
            if key in self._cells:
                c = self._cells[key]
                self._cells_mru.append(c)
                return c
            c = Cell(self._sheet.cell(key), self)
            self._cells[key] = c
            self._cells_mru.append(c)
            return c
        raise TypeError("Only string references (e.g., 'A1') are supported")

    def cell(self, row, column, value=None):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        key = (row, column)
        if key in self._cells:
            c = self._cells[key]
        else:
            c = Cell(self._sheet.cell(row, column), self)
            self._cells[key] = c

        if value is not None:
            c.value = value
        self._cells_mru.append(c)
        return c

    def _get_cached_cell(self, raw_cell):
        """Internal helper to get a cached Cell object from a raw XLCell."""
        ref = raw_cell.cell_reference()
        key = (ref.row(), ref.column())
        if key in self._cells:
            c = self._cells[key]
            self._cells_mru.append(c)
            return c
        c = Cell(raw_cell, self)
        self._cells[key] = c
        self._cells_mru.append(c)
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

    @property
    def zoom(self):
        """Get or set the worksheet zoom scale (percentage, e.g., 100)."""
        return self._sheet.zoom()

    @zoom.setter
    def zoom(self, value):
        self._sheet.set_zoom(int(value))

    def group_rows(self, row_first, row_last, outline_level=1, collapsed=False):
        self._sheet.group_rows(row_first, row_last, outline_level, collapsed)

    def group_columns(self, col_first, col_last, outline_level=1, collapsed=False):
        self._sheet.group_columns(col_first, col_last, outline_level, collapsed)

    def last_cell(self):
        return self._sheet.last_cell()

    def get_row(self, row_number: int):
        """Return the native XLRow handle for a 1-based row index."""
        return self._sheet.row(row_number)

    def iter_native_rows(self, first=None, last=None):
        """Iterate native XLRow objects (distinct from ``rows`` Cell tuples)."""
        if first is None and last is None:
            return self._sheet.rows()
        if last is None:
            return self._sheet.rows(first)
        return self._sheet.rows(first, last)

    def find_cell(self, ref_or_row, col=None):
        if col is None:
            return self._sheet.find_cell(ref_or_row)
        return self._sheet.find_cell(ref_or_row, col)

    def set_show_grid_lines(self, show: bool):
        self._sheet.set_show_grid_lines(show)

    def show_grid_lines(self) -> bool:
        return self._sheet.show_grid_lines()

    def stream_writer(self, use_shared_strings=False, max_unique_strings=100000):
        """
        Get a stream writer for this worksheet.

        Returns a :class:`~pyopenxlsx.streams.StreamWriter` that applies the
        same date/datetime coercion and ``auto_date_formats`` behaviour as
        bulk worksheet writes.

        :param use_shared_strings: When True, reuse shared-string table entries (saves space for
            repeated text at the cost of a bounded local cache).
        :param max_unique_strings: Cap on unique strings cached when use_shared_strings is True.
        """
        from .streams import StreamWriter

        native = self._sheet.stream_writer(use_shared_strings, max_unique_strings)
        return StreamWriter(native, self._workbook)

    def stream_reader(
        self, options=None, *, empty_rows=None, apply_number_formats=None
    ):
        """
        Get a stream reader for this worksheet.

        :param options: Optional XLStreamReadOptions instance.
        :param empty_rows: XLStreamEmptyRowPolicy (or pass via options).
        :param apply_number_formats: When True, format numeric cells as display strings where
            applicable (used by next_row_strings).
        """
        from .streams import StreamReader

        if options is None and (
            empty_rows is not None or apply_number_formats is not None
        ):
            from ._openxlsx import XLStreamReadOptions

            options = XLStreamReadOptions()
            if empty_rows is not None:
                options.empty_rows = empty_rows
            if apply_number_formats is not None:
                options.apply_number_formats = apply_number_formats
        if options is None:
            native = self._sheet.stream_reader()
        else:
            native = self._sheet.stream_reader(options)
        return StreamReader(native)

    def auto_fit_column(self, column_number: int):
        """Auto-fit the specified column."""
        self._sheet.auto_fit_column(column_number)
