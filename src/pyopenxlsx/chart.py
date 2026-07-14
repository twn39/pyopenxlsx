"""High-level chart helpers over the native XLChart surface.

Prefer ``Worksheet.add_chart`` for placement; use :class:`Chart` for a
fluent series/title/legend API. String chart type names are supported
(e.g. ``\"bar\"``, ``\"column\"``).
"""

from __future__ import annotations

from typing import Any, Optional, Sequence, Union

from ._openxlsx import XLChartType, XLLegendPosition


def chart_type(name: Union[str, XLChartType]) -> XLChartType:
    """Resolve a chart type from ``XLChartType`` or a case-insensitive name.

    Examples: ``\"bar\"``, ``\"Column\"``, ``\"line\"``, ``\"pie\"``.
    """
    if isinstance(name, XLChartType):
        return name
    key = name.strip().replace("-", "").replace("_", "").lower()
    for attr in dir(XLChartType):
        if attr.startswith("_"):
            continue
        if attr.replace("_", "").lower() == key:
            return getattr(XLChartType, attr)
    raise ValueError(f"Unknown chart type: {name!r}")


def legend_position(name: Union[str, XLLegendPosition]) -> XLLegendPosition:
    """Resolve legend position from enum or name (``right``, ``bottom``, …)."""
    if isinstance(name, XLLegendPosition):
        return name
    key = name.strip().replace("-", "").replace("_", "").lower()
    for attr in dir(XLLegendPosition):
        if attr.startswith("_"):
            continue
        if attr.replace("_", "").lower() == key:
            return getattr(XLLegendPosition, attr)
    raise ValueError(f"Unknown legend position: {name!r}")


class Chart:
    """Fluent wrapper around a native ``XLChart``.

    Does not own placement; obtain via :func:`add_chart` or wrap an existing
    chart with ``Chart(native)``.
    """

    __slots__ = ("_chart",)

    def __init__(self, native_chart: Any):
        self._chart = native_chart

    @property
    def raw(self) -> Any:
        """Underlying native chart object."""
        return self._chart

    def title(self, text: str) -> "Chart":
        self._chart.set_title(text)
        return self

    def style(self, style_id: int) -> "Chart":
        self._chart.set_style(style_id)
        return self

    def legend(self, position: Union[str, XLLegendPosition] = "right") -> "Chart":
        self._chart.set_legend_position(legend_position(position))
        return self

    def series(
        self,
        values_ref: str,
        *,
        name: str = "",
        categories_ref: str = "",
        series_type: Optional[Union[str, XLChartType]] = None,
        secondary_axis: bool = False,
    ) -> "Chart":
        """Add a series from A1-style references."""
        ct = None if series_type is None else chart_type(series_type)
        self._chart.add_series_ref(
            values_ref, name, categories_ref, ct, secondary_axis
        )
        return self

    def series_many(
        self,
        series_list: Sequence[Any],
        *,
        categories_ref: str = "",
    ) -> "Chart":
        """Add multiple series.

        Each item is either a values ref string, or a mapping/tuple::

            \"Sheet1!$B$2:$B$4\"
            (\"Sheet1!$B$2:$B$4\", \"Sales\")
            {\"values\": \"...\", \"name\": \"Sales\", \"categories\": \"...\"}
        """
        for item in series_list:
            if isinstance(item, str):
                self.series(item, categories_ref=categories_ref)
            elif isinstance(item, (tuple, list)):
                values = item[0]
                name = item[1] if len(item) > 1 else ""
                cats = item[2] if len(item) > 2 else categories_ref
                self.series(values, name=name, categories_ref=cats)
            elif isinstance(item, dict):
                self.series(
                    item["values"],
                    name=item.get("name", ""),
                    categories_ref=item.get("categories", categories_ref),
                    series_type=item.get("chart_type") or item.get("series_type"),
                    secondary_axis=bool(item.get("secondary_axis", False)),
                )
            else:
                raise TypeError(f"Unsupported series item: {type(item)!r}")
        return self

    def bubble_series(
        self,
        x_ref: str,
        y_ref: str,
        size_ref: str,
        *,
        name: str = "",
    ) -> "Chart":
        self._chart.add_bubble_series(x_ref, y_ref, size_ref, name)
        return self

    def data_labels(
        self,
        *,
        value: bool = True,
        category: bool = False,
        percent: bool = False,
    ) -> "Chart":
        self._chart.set_show_data_labels(value, category, percent)
        return self

    def data_table(self, show: bool = True, *, keys: bool = False) -> "Chart":
        self._chart.set_show_data_table(show, keys)
        return self

    def overlap(self, percent: int) -> "Chart":
        self._chart.set_overlap(percent)
        return self

    def gap_width(self, percent: int) -> "Chart":
        self._chart.set_gap_width(percent)
        return self

    def hole_size(self, percent: int) -> "Chart":
        self._chart.set_hole_size(percent)
        return self

    def rotation(self, x: int, y: int, perspective: int = 30) -> "Chart":
        self._chart.set_rotation(x, y, perspective)
        return self

    def chart_area_color(self, hex_rgb: str) -> "Chart":
        self._chart.set_chart_area_color(hex_rgb)
        return self

    def x_axis(self) -> Any:
        return self._chart.x_axis()

    def y_axis(self) -> Any:
        return self._chart.y_axis()

    def __getattr__(self, name: str) -> Any:
        return getattr(self._chart, name)


def add_chart(
    worksheet: Any,
    type_name: Union[str, XLChartType],
    name: str,
    *,
    row: int = 5,
    col: int = 5,
    width: int = 400,
    height: int = 300,
    title: Optional[str] = None,
    series_ref: Optional[str] = None,
    series_name: str = "",
    cats_ref: Optional[str] = None,
    series: Optional[Sequence[Any]] = None,
    legend: Optional[str] = None,
    wrap: bool = True,
) -> Any:
    """Create a chart with optional title, series, and legend.

    :param wrap: When True (default), return a :class:`Chart` wrapper;
        when False, return the native chart object (legacy behaviour).
    """
    ct = chart_type(type_name)
    native = worksheet._sheet.add_chart(ct, name, row, col, width, height)
    chart = Chart(native)
    if title is not None:
        chart.title(title)
    if series_ref is not None:
        chart.series(series_ref, name=series_name, categories_ref=cats_ref or "")
    if series:
        chart.series_many(series, categories_ref=cats_ref or "")
    if legend is not None:
        chart.legend(legend)
    return chart if wrap else native
