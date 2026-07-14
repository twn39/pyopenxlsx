from typing import Any, Optional

from pyopenxlsx._openxlsx import (
    XLCalculationEngine,
    XLCalculationOptions,
    XLEvalSession,
    XLFormulaDiagnosticReporter,
    XLFormulaEngine,
)


class FormulaEngine:
    """
    Lightweight formula evaluation engine.
    """

    def __init__(self):
        self._engine = XLFormulaEngine()

    def evaluate(
        self,
        formula: str,
        worksheet=None,
        *,
        session: Optional[XLEvalSession] = None,
        reporter: Optional[XLFormulaDiagnosticReporter] = None,
        current_row: Optional[int] = None,
        current_col: Optional[int] = None,
        current_sheet: Optional[str] = None,
    ) -> Any:
        """
        Evaluate a formula string.

        If a worksheet is provided, cell references within the formula will be resolved.
        Optionally pass session / reporter, or current_row/col/sheet for parameterless
        ROW()/COLUMN() and relative features.
        """
        wks_binding = worksheet._sheet if worksheet else None
        owned_session = None
        if session is None and (
            current_row is not None
            or current_col is not None
            or current_sheet is not None
        ):
            owned_session = XLEvalSession()
            if current_row is not None and current_col is not None:
                owned_session.set_current_cell(int(current_row), int(current_col))
            if current_sheet is not None:
                owned_session.set_current_sheet(current_sheet)
            session = owned_session
        return self._engine.evaluate(formula, wks_binding, session, reporter)


class CalculationEngine:
    """
    Sheet- or workbook-scoped formula recalculation with dependency tracking.
    """

    def __init__(self, target, options: Optional[XLCalculationOptions] = None):
        """
        :param target: A Worksheet or Workbook instance.
        :param options: Optional XLCalculationOptions.
        """
        if hasattr(target, "_sheet"):
            # Worksheet
            self._engine = XLCalculationEngine(target._sheet, options)
        elif hasattr(target, "_doc"):
            # Workbook
            self._engine = XLCalculationEngine(target._doc, options)
        else:
            # Raw XLWorksheet / XLDocument
            self._engine = XLCalculationEngine(target, options)

    def rebuild(self) -> None:
        self._engine.rebuild()

    @property
    def formula_count(self) -> int:
        return self._engine.formula_count

    @property
    def dirty_count(self) -> int:
        return self._engine.dirty_count

    def calc_cell_value(self, a1: str) -> Any:
        return self._engine.calc_cell_value(a1)

    def recalculate(self) -> int:
        return self._engine.recalculate()

    def recalculate_all(self) -> int:
        return self._engine.recalculate_all()

    def mark_dirty(self, a1: str, propagate: bool = True) -> None:
        self._engine.mark_dirty(a1, propagate)

    def set_input_value(self, a1: str, value: Any) -> None:
        self._engine.set_input_value(a1, value)
