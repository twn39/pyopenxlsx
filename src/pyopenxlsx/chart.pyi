from typing import Any, Optional, Sequence, Union
from ._openxlsx import XLChartType, XLLegendPosition

def chart_type(name: Union[str, XLChartType]) -> XLChartType: ...
def legend_position(name: Union[str, XLLegendPosition]) -> XLLegendPosition: ...

class Chart:
    def __init__(self, native_chart: Any) -> None: ...
    @property
    def raw(self) -> Any: ...
    def title(self, text: str) -> Chart: ...
    def style(self, style_id: int) -> Chart: ...
    def legend(self, position: Union[str, XLLegendPosition] = ...) -> Chart: ...
    def series(
        self,
        values_ref: str,
        *,
        name: str = ...,
        categories_ref: str = ...,
        series_type: Optional[Union[str, XLChartType]] = ...,
        secondary_axis: bool = ...,
    ) -> Chart: ...
    def series_many(
        self,
        series_list: Sequence[Any],
        *,
        categories_ref: str = ...,
    ) -> Chart: ...
    def bubble_series(
        self,
        x_ref: str,
        y_ref: str,
        size_ref: str,
        *,
        name: str = ...,
    ) -> Chart: ...
    def data_labels(
        self,
        *,
        value: bool = ...,
        category: bool = ...,
        percent: bool = ...,
    ) -> Chart: ...
    def data_table(self, show: bool = ..., *, keys: bool = ...) -> Chart: ...
    def overlap(self, percent: int) -> Chart: ...
    def gap_width(self, percent: int) -> Chart: ...
    def hole_size(self, percent: int) -> Chart: ...
    def rotation(self, x: int, y: int, perspective: int = ...) -> Chart: ...
    def chart_area_color(self, hex_rgb: str) -> Chart: ...
    def x_axis(self) -> Any: ...
    def y_axis(self) -> Any: ...
    def __getattr__(self, name: str) -> Any: ...

def add_chart(
    worksheet: Any,
    type_name: Union[str, XLChartType],
    name: str,
    *,
    row: int = ...,
    col: int = ...,
    width: int = ...,
    height: int = ...,
    title: Optional[str] = ...,
    series_ref: Optional[str] = ...,
    series_name: str = ...,
    cats_ref: Optional[str] = ...,
    series: Optional[Sequence[Any]] = ...,
    legend: Optional[str] = ...,
    wrap: bool = ...,
) -> Any: ...
