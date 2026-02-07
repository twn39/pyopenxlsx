from weakref import ref as weakref



class Column:
    """
    Represents an Excel column.

    Uses weak references to avoid circular references with Worksheet/Workbook.
    """

    __slots__ = ("_column", "_worksheet_ref", "_workbook_ref")

    def __init__(self, raw_column, worksheet=None):
        self._column = raw_column
        # Use weak references to avoid circular references
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
    def width(self):
        return self._column.width()

    @width.setter
    def width(self, value):
        self._column.set_width(float(value))

    @property
    def hidden(self):
        return self._column.is_hidden()

    @hidden.setter
    def hidden(self, value):
        self._column.set_hidden(bool(value))

    @property
    def style_index(self):
        return self._column.format()

    @style_index.setter
    def style_index(self, value):
        self._column.set_format(value)
