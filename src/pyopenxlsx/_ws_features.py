from typing import Any

from .data_validation import DataValidations
from .table import Table
from .autofilter import AutoFilter


class WorksheetFeaturesMixin:
    # Provided by Worksheet via mixin composition (for type checkers).
    _sheet: Any
    _workbook: Any
    _closed: bool
    max_row: int
    max_column: int

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

    def pivot_tables(self):
        return self._sheet.pivot_tables()

    def delete_pivot_table(self, name: str):
        return self._sheet.delete_pivot_table(name)

    def slicers(self):
        """Return the slicer collection for this worksheet."""
        return self._sheet.slicers()

    def delete_slicer(self, name: str):
        """Delete a slicer by name."""
        self._sheet.delete_slicer(name)

    def apply_auto_filter(self):
        """Apply auto filter to the worksheet."""
        self._sheet.apply_auto_filter()

    def add_conditional_formatting(self, sqref: str, rule):
        """Add conditional formatting to a range.

        Prefer builders from :mod:`pyopenxlsx.conditional_formatting`
        (e.g. ``color_scale``, ``data_bar``) or native ``XL*`` rule objects.
        """
        self._sheet.add_conditional_formatting(sqref, rule)

    def remove_conditional_formatting(self, sqref: str):
        """Remove conditional formatting from a range."""
        self._sheet.remove_conditional_formatting(sqref)

    def clear_all_conditional_formatting(self):
        """Clear all conditional formatting."""
        self._sheet.clear_all_conditional_formatting()

    def add_sparkline(
        self, location: str, data_range: str, sparkline_type=None, options=None
    ):
        """Add a sparkline to the worksheet."""
        if options is not None:
            self._sheet.add_sparkline(location, data_range, options)
        elif sparkline_type is not None:
            self._sheet.add_sparkline(location, data_range, sparkline_type)
        else:
            self._sheet.add_sparkline(location, data_range)

    def add_comment(self, cell_ref: str, text: str, author: str = ""):
        """Add a simple (legacy) comment."""
        self._sheet.add_comment(cell_ref, text, author)

    def add_threaded_comment(self, cell_ref: str, text: str, author: str = ""):
        """Add a modern threaded comment."""
        return self._sheet.add_threaded_comment(cell_ref, text, author)

    def add_threaded_reply(self, parent_id: str, text: str, author: str = ""):
        """Add a reply to a threaded comment."""
        return self._sheet.add_threaded_reply(parent_id, text, author)

    def add_chart(
        self,
        chart_type,
        name: str,
        row: int = 5,
        col: int = 5,
        width: int = 400,
        height: int = 300,
        *,
        wrap: bool = True,
    ):
        """
        Add a chart to the worksheet (high-level facade over the native sheet).

        ``chart_type`` may be an ``XLChartType`` or a friendly name such as
        ``\"bar\"`` / ``\"column\"`` (see :func:`pyopenxlsx.chart.chart_type`).

        Returns a :class:`~pyopenxlsx.chart.Chart` wrapper when *wrap* is True
        (default); set ``wrap=False`` for the raw native object.
        """
        from .chart import Chart, chart_type as resolve_chart_type

        resolved = resolve_chart_type(chart_type)
        native = self._sheet.add_chart(resolved, name, row, col, width, height)
        return Chart(native) if wrap else native

    def add_chart_anchor(self, chart_type, anchor, *, wrap: bool = True):
        """Add a chart using an ``XLChartAnchor`` (high-level facade)."""
        from .chart import Chart, chart_type as resolve_chart_type

        resolved = resolve_chart_type(chart_type)
        native = self._sheet.add_chart_anchor(resolved, anchor)
        return Chart(native) if wrap else native

    def add_pivot_table(self, options):
        """
        Add a pivot table from ``XLPivotTableOptions`` or a
        :class:`~pyopenxlsx.pivot.PivotTableBuilder`.
        """
        from .pivot import PivotTableBuilder

        if isinstance(options, PivotTableBuilder):
            options = options.options
        return self._sheet.add_pivot_table(options)
