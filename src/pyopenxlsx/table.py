from weakref import ref as weakref


class Table:
    """
    Represents an Excel Table (ListObject).
    """

    __slots__ = ("_table", "_worksheet_ref")

    def __init__(self, raw_table, worksheet=None):
        self._table = raw_table
        self._worksheet_ref = weakref(worksheet) if worksheet else None

    @property
    def _worksheet(self):
        return self._worksheet_ref() if self._worksheet_ref else None

    @property
    def name(self):
        """The name of the table. Table names cannot have spaces."""
        return self._table.name()

    @name.setter
    def name(self, value):
        self._table.set_name(str(value))

    @property
    def display_name(self):
        """The display name of the table."""
        return self._table.display_name()

    @display_name.setter
    def display_name(self, value):
        self._table.set_display_name(str(value))

    @property
    def range(self):
        """The range reference of the table (e.g., 'A1:C10')."""
        return self._table.range_reference()

    @range.setter
    def range(self, value):
        self._table.set_range_reference(str(value))

    @property
    def style(self):
        """The table style name (e.g., 'TableStyleMedium2')."""
        return self._table.style_name()

    @style.setter
    def style(self, value):
        self._table.set_style_name(str(value))

    @property
    def show_row_stripes(self):
        """Whether row stripes are shown."""
        return self._table.show_row_stripes()

    @show_row_stripes.setter
    def show_row_stripes(self, value):
        self._table.set_show_row_stripes(bool(value))

    @property
    def show_column_stripes(self):
        """Whether column stripes are shown."""
        return self._table.show_column_stripes()

    @show_column_stripes.setter
    def show_column_stripes(self, value):
        self._table.set_show_column_stripes(bool(value))

    @property
    def show_first_column(self):
        """Whether the first column is highlighted."""
        return self._table.show_first_column()

    @show_first_column.setter
    def show_first_column(self, value):
        self._table.set_show_first_column(bool(value))

    @property
    def show_last_column(self):
        """Whether the last column is highlighted."""
        return self._table.show_last_column()

    @show_last_column.setter
    def show_last_column(self, value):
        self._table.set_show_last_column(bool(value))

    @property
    def show_totals_row(self):
        """Whether the totals row is shown."""
        return self._table.show_totals_row()

    @show_totals_row.setter
    def show_totals_row(self, value):
        self._table.set_show_totals_row(bool(value))

    def append_column(self, name):
        """Append a new column to the table."""
        self._table.append_column(str(name))
