from typing import Any, Optional, Union

from pyopenxlsx.worksheet import Worksheet
from pyopenxlsx.workbook import Workbook
from pyopenxlsx._openxlsx import (
    XLCalculationOptions,
    XLEvalSession,
    XLFormulaDiagnosticReporter,
)

class FormulaEngine:
    def __init__(self) -> None: ...
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

class CalculationEngine:
    def __init__(
        self,
        target: Union[Worksheet, Workbook, Any],
        options: Optional[XLCalculationOptions] = None,
    ) -> None: ...
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
