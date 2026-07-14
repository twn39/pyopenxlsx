from . import _coercion
from .formula import Formula

# Re-export for public / historical import paths (``from pyopenxlsx.cell import …``).
datetime_to_serial = _coercion.datetime_to_serial
serial_to_datetime = _coercion.serial_to_datetime


class Cell:
    """
    Represents an Excel cell.

    OPTIMIZATION PRINCIPLE:
    Direct reference model is used for `_worksheet_val` instead of `weakref`.
    Because the Worksheet keeps a WeakValueDictionary `_cells` cache of Cell objects,
    there is no strong reference cycle between Worksheet and Cell.
    This eliminates the allocation of 2 `weakref` objects for every Cell creation,
    massively improving performance in loops and preventing GC thrashing.
    """

    # Include __weakref__ to allow weak references to Cell objects
    # This enables WeakValueDictionary caching in Worksheet
    __slots__ = ("_cell", "_worksheet_val", "__weakref__")

    def __init__(self, raw_cell, worksheet=None):
        self._cell = raw_cell
        self._worksheet_val = worksheet

    @property
    def _closed(self):
        """Return True if the parent worksheet/workbook has been closed."""
        return self._worksheet_val._closed if self._worksheet_val else False

    @property
    def _worksheet(self):
        """Get the worksheet, or None if not set."""
        return self._worksheet_val

    @property
    def _workbook(self):
        """Get the workbook, or None if not set."""
        return self._worksheet_val._workbook if self._worksheet_val else None

    @property
    def comment(self):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        if self._worksheet is None:
            return None
        text = self._worksheet._sheet.comments().get(
            self._cell.cell_reference().address()
        )
        if not text:
            return None
        return text

    @comment.setter
    def comment(self, value):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        if self._worksheet is None:
            raise ValueError("Cell must be associated with a worksheet to set comments")

        comments = self._worksheet._sheet.comments()
        # Ensure at least one author exists, otherwise Excel may report the file as corrupt
        if comments.author_count() == 0:
            comments.add_author("pyopenxlsx")

        addr = self._cell.cell_reference().address()
        if value is None:
            comments.delete_comment(addr)
        else:
            val_str = str(value)

            # --- Auto-size Logic ---
            try:
                # Estimate dimensions based on text content
                lines = val_str.split("\n")
                line_count = len(lines)

                # More accurate width estimation: count chars, giving more weight to non-ASCII
                def content_width(s):
                    return sum(2 if ord(c) > 127 else 1 for c in s)

                max_width = max(content_width(line) for line in lines) if lines else 0

                # Heuristic settings:
                # Average column width handles ~10-12 units of content_width
                col_span = max(3, int(max_width / 10) + 1)
                if col_span > 12:
                    col_span = 12  # Moderate cap for width

                # Average row height handles 1 line of text
                row_span = line_count + 1
                if row_span < 3:
                    row_span = 3  # Minimum height

                comments.set(addr, val_str, 0, col_span, row_span)
            except Exception:
                comments.set(addr, val_str)

    @property
    def value(self):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        val = self._cell.value
        if isinstance(val, (float, int)) and self.is_date:
            try:
                return serial_to_datetime(val)
            except Exception:
                pass
        return val

    @value.setter
    def value(self, val):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        storage, date_kind = _coercion.coerce_cell_value(val)
        self._cell.value = storage
        if date_kind is None:
            return
        wb = self._workbook
        if not _coercion.workbook_wants_auto_date(wb):
            return
        if self.is_date:
            return
        self.style_index = wb._get_auto_date_style(
            is_datetime=(date_kind == _coercion.DATE_KIND_DATETIME)
        )

    @property
    def formula(self):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        return Formula(self._cell)

    @formula.setter
    def formula(self, val):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        self._cell.set_formula(str(val))

    @property
    def style_index(self):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        return self._cell.cell_format()

    @style_index.setter
    def style_index(self, val):
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        self._cell.set_cell_format(val)

    @property
    def style(self):
        return self.style_index

    @style.setter
    def style(self, val):
        if hasattr(val, "style_index"):
            self.style_index = val.style_index
        else:
            self.style_index = val

    @property
    def font(self):
        if self._workbook is None:
            return None
        cf = self._workbook.styles.cell_formats().cell_format_by_index(self.style_index)
        return self._workbook.styles.fonts().font_by_index(cf.font_index())

    @property
    def fill(self):
        if self._workbook is None:
            return None
        cf = self._workbook.styles.cell_formats().cell_format_by_index(self.style_index)
        return self._workbook.styles.fills().fill_by_index(cf.fill_index())

    @property
    def border(self):
        if self._workbook is None:
            return None
        cf = self._workbook.styles.cell_formats().cell_format_by_index(self.style_index)
        return self._workbook.styles.borders().border_by_index(cf.border_index())

    @property
    def alignment(self):
        if self._workbook is None:
            return None
        cf = self._workbook.styles.cell_formats().cell_format_by_index(self.style_index)
        return cf.alignment()

    @property
    def is_date(self):
        """
        Returns True if the cell is formatted as a date/time.
        Requires workbook to be passed to Cell constructor.
        """
        if self._closed is True:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        return _coercion.style_is_date_format(self._workbook, self.style_index)
