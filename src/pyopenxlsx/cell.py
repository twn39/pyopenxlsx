from .formula import Formula
from .styles import is_date_format
from datetime import datetime, date, timedelta
from weakref import ref as weakref


def datetime_to_serial(val):
    if isinstance(val, date) and not isinstance(val, datetime):
        val = datetime.combine(val, datetime.min.time())
    delta = val - datetime(1899, 12, 30)
    return delta.total_seconds() / 86400.0


def serial_to_datetime(serial):
    return datetime(1899, 12, 30) + timedelta(days=serial)


class Cell:
    """
    Represents an Excel cell.

    Uses weak references to worksheet and workbook to avoid circular references
    that could delay garbage collection.
    """

    # Include __weakref__ to allow weak references to Cell objects
    # This enables WeakValueDictionary caching in Worksheet
    __slots__ = ("_cell", "_worksheet_ref", "_workbook_ref", "__weakref__")

    def __init__(self, raw_cell, worksheet=None):
        self._cell = raw_cell
        # Use weak references to avoid circular references with Worksheet
        self._worksheet_ref = weakref(worksheet) if worksheet else None
        self._workbook_ref = (
            weakref(worksheet._workbook) if worksheet and worksheet._workbook else None
        )

    @property
    def _worksheet(self):
        """Get the worksheet, or None if it has been garbage collected."""
        return self._worksheet_ref() if self._worksheet_ref else None

    @property
    def _workbook(self):
        """Get the workbook, or None if it has been garbage collected."""
        return self._workbook_ref() if self._workbook_ref else None

    @property
    def comment(self):
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
        if self._worksheet is None:
            raise ValueError("Cell must be associated with a worksheet to set comments")
        addr = self._cell.cell_reference().address()
        if value is None:
            self._worksheet._sheet.comments().delete_comment(addr)
        else:
            self._worksheet._sheet.comments().set(addr, str(value))

    @property
    def value(self):
        val = self._cell.value
        if isinstance(val, (float, int)) and self.is_date:
            try:
                return serial_to_datetime(val)
            except Exception:
                pass
        return val

    @value.setter
    def value(self, val):
        if isinstance(val, (date, datetime)):
            val = datetime_to_serial(val)
        self._cell.value = val

    @property
    def formula(self):
        return Formula(self._cell)

    @formula.setter
    def formula(self, val):
        self._cell.set_formula(str(val))

    @property
    def style_index(self):
        return self._cell.cell_format()

    @style_index.setter
    def style_index(self, val):
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
        if self._workbook is None:
            return False

        style_idx = self.style_index
        if style_idx < 0:
            return False

        # Check cache
        if style_idx in self._workbook._date_format_cache:
            return self._workbook._date_format_cache[style_idx]

        # Get styles from workbook
        styles = self._workbook.styles
        cfs = styles.cell_formats()
        if style_idx >= cfs.count():
            self._workbook._date_format_cache[style_idx] = False
            return False

        cf = cfs.cell_format_by_index(style_idx)
        nf_id = cf.number_format_id()

        # Check standard formats
        if is_date_format(nf_id):
            self._workbook._date_format_cache[style_idx] = True
            return True

        # Check custom formats via string
        nfs = styles.number_formats()
        try:
            val = nfs.number_format_by_id(nf_id)
            if val:
                res = is_date_format(val.format_code())
                self._workbook._date_format_cache[style_idx] = res
                return res
        except Exception:
            pass

        self._workbook._date_format_cache[style_idx] = False
        return False
