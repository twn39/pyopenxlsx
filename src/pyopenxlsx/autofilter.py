from weakref import ref as weakref
from ._openxlsx import XLFilterLogic


class FilterColumn:
    """
    Represents a specific column in an AutoFilter range.
    """

    def __init__(self, raw_column, autofilter=None):
        self._column = raw_column
        self._autofilter_ref = weakref(autofilter) if autofilter else None

    @property
    def col_id(self):
        """0-based column ID relative to the AutoFilter range."""
        return self._column.col_id()

    def add_filter(self, value):
        """Add a specific value to filter by."""
        self._column.add_filter(str(value))

    def clear(self):
        """Clear all filters for this column."""
        self._column.clear_filters()

    def set_custom_filter(self, op, val, logic=None, op2=None, val2=None):
        """
        Set custom filter criteria.

        :param op: Comparison operator (e.g., 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual')
        :param val: Value to compare against.
        :param logic: Logical operator ('and', 'or') if using a compound filter.
        :param op2: Second comparison operator.
        :param val2: Second value to compare against.
        """
        if logic is None:
            self._column.set_custom_filter(str(op), str(val))
        else:
            logic_enum = XLFilterLogic.And if logic.lower() == "and" else XLFilterLogic.Or
            self._column.set_custom_filter(str(op), str(val), logic_enum, str(op2), str(val2))

    def set_top10(self, value, percent=False, top=True):
        """
        Set a top-10 filter.

        :param value: Threshold value.
        :param percent: If True, filters by top percentage rather than count.
        :param top: If True, filters top values; if False, filters bottom values.
        """
        self._column.set_top10(float(value), bool(percent), bool(top))


class AutoFilter:
    """
    Represents an Excel AutoFilter.
    """

    def __init__(self, raw_autofilter, worksheet=None):
        self._autofilter = raw_autofilter
        self._worksheet_ref = weakref(worksheet) if worksheet else None

    def __bool__(self):
        return bool(self._autofilter)

    @property
    def ref(self):
        """The reference range of the AutoFilter (e.g., 'A1:C10')."""
        return self._autofilter.ref()

    @ref.setter
    def ref(self, value):
        self._autofilter.set_ref(str(value))

    def filter_column(self, col_id):
        """
        Get or create a filter column by its 0-based ID relative to the range.

        :param col_id: 0-based column ID.
        :return: FilterColumn object.
        """
        return FilterColumn(self._autofilter.filter_column(col_id), self)

    def __getitem__(self, col_id):
        return self.filter_column(col_id)
    def __eq__(self, other):
        if isinstance(other, str):
            return self.ref == other
        if isinstance(other, AutoFilter):
            return self.ref == other.ref
        return False

    def __str__(self):
        return self.ref
