"""Fluent builder for pivot tables over ``XLPivotTableOptions``."""

from __future__ import annotations

from typing import Any, List, Optional, Sequence, Union

from ._openxlsx import XLPivotSubtotal, XLPivotTableOptions


def pivot_subtotal(name: Union[str, XLPivotSubtotal]) -> XLPivotSubtotal:
    """Resolve a subtotal from enum or friendly name (``sum``, ``count``, …)."""
    if isinstance(name, XLPivotSubtotal):
        return name
    key = name.strip().replace("-", "").replace("_", "").lower()
    for attr in dir(XLPivotSubtotal):
        if attr.startswith("_"):
            continue
        if attr.replace("_", "").lower() == key:
            return getattr(XLPivotSubtotal, attr)
    raise ValueError(f"Unknown pivot subtotal: {name!r}")


class PivotTableBuilder:
    """Build and attach a pivot table with a chainable API.

    Example::

        PivotTableBuilder("SalesPivot", "Sheet1!A1:C20", "E1") \\
            .rows("Region") \\
            .columns("Year") \\
            .data("Amount", subtotal="sum") \\
            .style("PivotStyleMedium9") \\
            .add_to(ws)
    """

    def __init__(
        self,
        name: str,
        source_range: str,
        target_cell: str = "A1",
    ):
        self._opts = XLPivotTableOptions(name, source_range, target_cell)

    @property
    def options(self) -> XLPivotTableOptions:
        """Underlying native options object."""
        return self._opts

    def rows(
        self,
        *fields: str,
        selected_items: Optional[Sequence[str]] = None,
    ) -> "PivotTableBuilder":
        items = list(selected_items) if selected_items is not None else []
        for field in fields:
            self._opts.add_row_field(field, items)
        return self

    def columns(
        self,
        *fields: str,
        selected_items: Optional[Sequence[str]] = None,
    ) -> "PivotTableBuilder":
        items = list(selected_items) if selected_items is not None else []
        for field in fields:
            self._opts.add_column_field(field, items)
        return self

    def filters(
        self,
        *fields: str,
        selected_items: Optional[Sequence[str]] = None,
    ) -> "PivotTableBuilder":
        items = list(selected_items) if selected_items is not None else []
        for field in fields:
            self._opts.add_filter_field(field, items)
        return self

    def data(
        self,
        field: str,
        *,
        name: str = "",
        subtotal: Union[str, XLPivotSubtotal] = "sum",
        num_fmt_id: int = 0,
    ) -> "PivotTableBuilder":
        self._opts.add_data_field(
            field, name, pivot_subtotal(subtotal), num_fmt_id
        )
        return self

    def style(self, style_name: str) -> "PivotTableBuilder":
        self._opts.set_pivot_table_style(style_name)
        return self

    def data_on_rows(self, value: bool = True) -> "PivotTableBuilder":
        self._opts.set_data_on_rows(value)
        return self

    def compact(self, value: bool = True) -> "PivotTableBuilder":
        self._opts.set_compact_data(value)
        return self

    def grand_totals(
        self, *, rows: Optional[bool] = None, columns: Optional[bool] = None
    ) -> "PivotTableBuilder":
        if rows is not None:
            self._opts.set_row_grand_totals(rows)
        if columns is not None:
            self._opts.set_col_grand_totals(columns)
        return self

    def show_headers(
        self, *, rows: Optional[bool] = None, columns: Optional[bool] = None
    ) -> "PivotTableBuilder":
        if rows is not None:
            self._opts.set_show_row_headers(rows)
        if columns is not None:
            self._opts.set_show_col_headers(columns)
        return self

    def stripes(
        self, *, rows: Optional[bool] = None, columns: Optional[bool] = None
    ) -> "PivotTableBuilder":
        if rows is not None:
            self._opts.set_show_row_stripes(rows)
        if columns is not None:
            self._opts.set_show_col_stripes(columns)
        return self

    def configure(self, **flags: bool) -> "PivotTableBuilder":
        """Set common boolean options by keyword.

        Supported keys map to ``set_*`` methods without the ``set_`` prefix,
        e.g. ``show_drill=True``, ``use_auto_formatting=True``.
        """
        for key, value in flags.items():
            method = getattr(self._opts, f"set_{key}", None)
            if method is None:
                raise ValueError(f"Unknown pivot option: {key!r}")
            method(value)
        return self

    def add_to(self, worksheet: Any) -> Any:
        """Create the pivot table on *worksheet* and return the native result."""
        return worksheet.add_pivot_table(self._opts)


def pivot_table(
    name: str,
    source_range: str,
    target_cell: str = "A1",
    *,
    rows: Optional[Union[str, Sequence[str]]] = None,
    columns: Optional[Union[str, Sequence[str]]] = None,
    data: Optional[Union[str, Sequence[str]]] = None,
    filters: Optional[Union[str, Sequence[str]]] = None,
    style: Optional[str] = None,
    subtotal: Union[str, XLPivotSubtotal] = "sum",
) -> PivotTableBuilder:
    """Convenience constructor with common fields pre-filled."""

    def _as_list(v: Optional[Union[str, Sequence[str]]]) -> List[str]:
        if v is None:
            return []
        if isinstance(v, str):
            return [v]
        return list(v)

    b = PivotTableBuilder(name, source_range, target_cell)
    if rows:
        b.rows(*_as_list(rows))
    if columns:
        b.columns(*_as_list(columns))
    if filters:
        b.filters(*_as_list(filters))
    for field in _as_list(data):
        b.data(field, subtotal=subtotal)
    if style:
        b.style(style)
    return b
