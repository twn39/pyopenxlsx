"""High-level formula evaluation and sheet recalculation façades."""

from __future__ import annotations

from typing import Any, List, Optional, Sequence, Union

from pyopenxlsx._openxlsx import (
    XLCalculationEngine,
    XLCalculationOptions,
    XLEvalSession,
    XLFormulaDiagnosticReporter,
    XLFormulaEngine,
)


def calculation_options(
    *,
    write_back: Optional[bool] = None,
    use_defined_names: Optional[bool] = None,
    **extra: Any,
) -> XLCalculationOptions:
    """Build ``XLCalculationOptions`` from keyword flags.

    Unknown keys are applied as ``setattr`` on the options object when present.
    """
    opts = XLCalculationOptions()
    if write_back is not None:
        opts.write_back = write_back
    if use_defined_names is not None:
        opts.use_defined_names = use_defined_names
    for key, value in extra.items():
        if hasattr(opts, key):
            setattr(opts, key, value)
        else:
            raise ValueError(f"Unknown calculation option: {key!r}")
    return opts


class FormulaEngine:
    """
    Lightweight formula evaluation engine (single expression).

    Prefer this for ad-hoc evaluation. For workbook/sheet recalculation of
    stored cell formulas, use :class:`CalculationEngine`.
    """

    def __init__(self):
        self._engine = XLFormulaEngine()

    @property
    def raw(self) -> XLFormulaEngine:
        """Underlying native engine."""
        return self._engine

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

        If a worksheet is provided, cell references within the formula will be
        resolved. Optionally pass session / reporter, or current_row/col/sheet
        for parameterless ``ROW()``/``COLUMN()`` and relative features.

        The formula may be written with or without a leading ``=``.
        """
        text = formula.strip()
        if text.startswith("="):
            text = text[1:]

        wks_binding = worksheet._sheet if worksheet is not None and hasattr(worksheet, "_sheet") else worksheet
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
        return self._engine.evaluate(text, wks_binding, session, reporter)

    def evaluate_many(
        self,
        formulas: Sequence[str],
        worksheet=None,
        **kwargs: Any,
    ) -> List[Any]:
        """Evaluate multiple formulas with the same worksheet/session kwargs."""
        return [self.evaluate(f, worksheet, **kwargs) for f in formulas]

    def sum(self, range_or_args: str, worksheet=None, **kwargs: Any) -> Any:
        """Evaluate ``SUM(...)`` over a range expression or argument list string."""
        return self.evaluate(f"SUM({range_or_args})", worksheet, **kwargs)

    def average(self, range_or_args: str, worksheet=None, **kwargs: Any) -> Any:
        """Evaluate ``AVERAGE(...)``."""
        return self.evaluate(f"AVERAGE({range_or_args})", worksheet, **kwargs)

    def count(self, range_or_args: str, worksheet=None, **kwargs: Any) -> Any:
        """Evaluate ``COUNT(...)``."""
        return self.evaluate(f"COUNT({range_or_args})", worksheet, **kwargs)

    def if_(
        self,
        condition: str,
        true_value: str,
        false_value: str = "FALSE",
        worksheet=None,
        **kwargs: Any,
    ) -> Any:
        """Evaluate ``IF(condition, true, false)``."""
        return self.evaluate(
            f"IF({condition}, {true_value}, {false_value})",
            worksheet,
            **kwargs,
        )


class CalculationEngine:
    """
    Sheet- or workbook-scoped formula recalculation with dependency tracking.
    """

    def __init__(
        self,
        target,
        options: Optional[Union[XLCalculationOptions, dict]] = None,
    ):
        """
        :param target: A Worksheet or Workbook instance (or raw XL types).
        :param options: ``XLCalculationOptions``, a dict of flags, or ``None``.
        """
        if isinstance(options, dict):
            options = calculation_options(**options)

        if hasattr(target, "_sheet"):
            self._engine = XLCalculationEngine(target._sheet, options)
        elif hasattr(target, "_doc"):
            self._engine = XLCalculationEngine(target._doc, options)
        else:
            self._engine = XLCalculationEngine(target, options)

    @property
    def raw(self) -> XLCalculationEngine:
        return self._engine

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

    def recalculate_inputs(self, values: dict) -> int:
        """Set multiple input cells then recalculate.

        :param values: Mapping of A1 address → value
        :return: Number of cells recalculated (engine-dependent).
        """
        for a1, value in values.items():
            self.set_input_value(a1, value)
        return self.recalculate()
