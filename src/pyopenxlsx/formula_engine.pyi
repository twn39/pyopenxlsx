from typing import Any, Dict, List, Optional, Sequence, Union

from pyopenxlsx.worksheet import Worksheet
from pyopenxlsx.workbook import Workbook
from pyopenxlsx._openxlsx import (
    XLCalculationEngine,
    XLCalculationOptions,
    XLEvalSession,
    XLFormulaDiagnosticReporter,
    XLFormulaEngine,
)

def calculation_options(
    *,
    write_back: Optional[bool] = ...,
    use_defined_names: Optional[bool] = ...,
    **extra: Any,
) -> XLCalculationOptions: ...

class FormulaEngine:
    def __init__(self) -> None: ...
    @property
    def raw(self) -> XLFormulaEngine: ...
    def evaluate(
        self,
        formula: str,
        worksheet: Optional[Worksheet] = None,
        *,
        session: Optional[XLEvalSession] = None,
        reporter: Optional[XLFormulaDiagnosticReporter] = None,
        current_row: Optional[int] = None,
        current_col: Optional[int] = None,
        current_sheet: Optional[str] = None,
    ) -> Any: ...
    def evaluate_many(
        self,
        formulas: Sequence[str],
        worksheet: Optional[Worksheet] = None,
        **kwargs: Any,
    ) -> List[Any]: ...
    def sum(
        self, range_or_args: str, worksheet: Optional[Worksheet] = None, **kwargs: Any
    ) -> Any: ...
    def average(
        self, range_or_args: str, worksheet: Optional[Worksheet] = None, **kwargs: Any
    ) -> Any: ...
    def count(
        self, range_or_args: str, worksheet: Optional[Worksheet] = None, **kwargs: Any
    ) -> Any: ...
    def if_(
        self,
        condition: str,
        true_value: str,
        false_value: str = ...,
        worksheet: Optional[Worksheet] = None,
        **kwargs: Any,
    ) -> Any: ...

class CalculationEngine:
    def __init__(
        self,
        target: Union[Worksheet, Workbook, Any],
        options: Optional[Union[XLCalculationOptions, Dict[str, Any]]] = None,
    ) -> None: ...
    @property
    def raw(self) -> XLCalculationEngine: ...
    def rebuild(self) -> None: ...
    @property
    def formula_count(self) -> int: ...
    @property
    def dirty_count(self) -> int: ...
    def calc_cell_value(self, a1: str) -> Any: ...
    def recalculate(self) -> int: ...
    def recalculate_all(self) -> int: ...
    def mark_dirty(self, a1: str, propagate: bool = True) -> None: ...
    def set_input_value(self, a1: str, value: Any) -> None: ...
    def recalculate_inputs(self, values: Dict[str, Any]) -> int: ...
