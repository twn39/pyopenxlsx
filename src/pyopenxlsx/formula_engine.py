from typing import Any
from pyopenxlsx._openxlsx import XLFormulaEngine


class FormulaEngine:
    """
    Lightweight formula evaluation engine.
    """

    def __init__(self):
        self._engine = XLFormulaEngine()

    def evaluate(self, formula: str, worksheet=None) -> Any:
        """
        Evaluate a formula string.
        If a worksheet is provided, cell references within the formula will be resolved.
        """
        wks_binding = worksheet._sheet if worksheet else None
        return self._engine.evaluate(formula, wks_binding)
