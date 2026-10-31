from typing import Any, Optional, Sequence, Union
from ._openxlsx import (
    XLPivotShowValuesAs,
    XLPivotShowValuesAsOptions,
    XLPivotSubtotal,
    XLPivotTableOptions,
)

def pivot_subtotal(name: Union[str, XLPivotSubtotal]) -> XLPivotSubtotal: ...
def pivot_show_values_as(
    val: Union[
        str, XLPivotShowValuesAs, XLPivotShowValuesAsOptions, tuple, dict, None
    ] = None,
) -> XLPivotShowValuesAsOptions: ...

class PivotTableBuilder:
    def __init__(
        self, name: str, source_range: str, target_cell: str = ...
    ) -> None: ...
    @property
    def options(self) -> XLPivotTableOptions: ...
    def rows(
        self,
        *fields: str,
        selected_items: Optional[Sequence[str]] = ...,
        compact: bool = ...,
        outline: bool = ...,
        show_all: bool = ...,
        insert_blank_row: bool = ...,
        default_subtotal: bool = ...,
    ) -> PivotTableBuilder: ...
    def columns(
        self,
        *fields: str,
        selected_items: Optional[Sequence[str]] = ...,
        compact: bool = ...,
        outline: bool = ...,
        show_all: bool = ...,
        insert_blank_row: bool = ...,
        default_subtotal: bool = ...,
    ) -> PivotTableBuilder: ...
    def filters(
        self, *fields: str, selected_items: Optional[Sequence[str]] = ...
    ) -> PivotTableBuilder: ...
    def data(
        self,
        field: str,
        *,
        name: str = ...,
        subtotal: Union[str, XLPivotSubtotal] = ...,
        num_fmt_id: int = ...,
        show_values_as: Optional[
            Union[str, XLPivotShowValuesAs, XLPivotShowValuesAsOptions, tuple, dict]
        ] = ...,
    ) -> PivotTableBuilder: ...
    def style(self, style_name: str) -> PivotTableBuilder: ...
    def data_on_rows(self, value: bool = ...) -> PivotTableBuilder: ...
    def compact(self, value: bool = ...) -> PivotTableBuilder: ...
    def grand_totals(
        self, *, rows: Optional[bool] = ..., columns: Optional[bool] = ...
    ) -> PivotTableBuilder: ...
    def show_headers(
        self, *, rows: Optional[bool] = ..., columns: Optional[bool] = ...
    ) -> PivotTableBuilder: ...
    def stripes(
        self, *, rows: Optional[bool] = ..., columns: Optional[bool] = ...
    ) -> PivotTableBuilder: ...
    def configure(self, **flags: bool) -> PivotTableBuilder: ...
    def add_to(self, worksheet: Any) -> Any: ...

def pivot_table(
    name: str,
    source_range: str,
    target_cell: str = ...,
    *,
    rows: Optional[Union[str, Sequence[str]]] = ...,
    columns: Optional[Union[str, Sequence[str]]] = ...,
    data: Optional[Union[str, Sequence[str]]] = ...,
    filters: Optional[Union[str, Sequence[str]]] = ...,
    style: Optional[str] = ...,
    subtotal: Union[str, XLPivotSubtotal] = ...,
    show_values_as: Optional[
        Union[str, XLPivotShowValuesAs, XLPivotShowValuesAsOptions, tuple, dict]
    ] = ...,
) -> PivotTableBuilder: ...
