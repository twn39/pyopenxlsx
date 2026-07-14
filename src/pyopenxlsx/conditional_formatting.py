"""Pythonic builders for conditional formatting rules.

These wrap the native ``XL*`` rule factories with simpler call sites
(RGB tuples, keyword names) while remaining compatible with
``Worksheet.add_conditional_formatting``.
"""

from __future__ import annotations

from typing import Any, Optional, Sequence, Tuple, Union

from ._openxlsx import (
    XLAboveAverageRule,
    XLCellIsRule,
    XLCfOperator,
    XLColor,
    XLColorScaleRule,
    XLContainsBlanksRule,
    XLContainsErrorsRule,
    XLContainsTextRule,
    XLDataBarRule,
    XLDuplicateValuesRule,
    XLFormulaRule,
    XLIconSetRule,
    XLNotContainsBlanksRule,
    XLNotContainsErrorsRule,
    XLNotContainsTextRule,
    XLTop10Rule,
)

ColorLike = Union[XLColor, Tuple[int, int, int], Sequence[int], str]


def _to_xl_color(color: ColorLike) -> XLColor:
    if isinstance(color, XLColor):
        return color
    if isinstance(color, str):
        # Accept hex like "#FF0000" or "FF0000"
        h = color.lstrip("#")
        if len(h) != 6:
            raise ValueError(f"Expected 6-digit hex colour, got {color!r}")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return XLColor(r, g, b)
    if len(color) != 3:
        raise ValueError("Colour must be RGB triple or XLColor")
    return XLColor(int(color[0]), int(color[1]), int(color[2]))


def color_scale(
    start: ColorLike,
    end: ColorLike,
    mid: Optional[ColorLike] = None,
) -> Any:
    """2- or 3-stop colour scale rule."""
    c1, c2 = _to_xl_color(start), _to_xl_color(end)
    if mid is None:
        return XLColorScaleRule(c1, c2)
    return XLColorScaleRule(c1, _to_xl_color(mid), c2)


def data_bar(color: ColorLike = (0, 0, 255), *, show_value: bool = True) -> Any:
    """Data-bar conditional formatting rule."""
    return XLDataBarRule(_to_xl_color(color), show_value)


def cell_is(operator: Union[str, XLCfOperator], formula: str, formula2: Optional[str] = None) -> Any:
    """Cell value comparison rule.

    ``operator`` may be an ``XLCfOperator`` or a short name such as
    ``\">\"``, ``\">=\"``, ``\"<\"``, ``\"<=\"``, ``\"=\"``, ``\"!=\"``,
    ``\"between\"``, ``\"notBetween\"``.
    """
    op = _parse_cf_operator(operator)
    if formula2 is not None:
        # Two-formula form uses string operator overload in native API.
        return XLCellIsRule(formula, formula2)
    return XLCellIsRule(op, formula)


def formula_rule(formula: str) -> Any:
    """Expression-based rule (true when formula is non-zero)."""
    return XLFormulaRule(formula)


def icon_set(
    name: str = "3TrafficLights1",
    *,
    show_value: bool = True,
    reverse: bool = False,
) -> Any:
    return XLIconSetRule(name, show_value, reverse)


def top10(rank: int = 10, *, percent: bool = False, bottom: bool = False) -> Any:
    return XLTop10Rule(rank, percent, bottom)


def above_average(
    *,
    above: bool = True,
    equal_average: bool = False,
    std_dev: int = 0,
) -> Any:
    return XLAboveAverageRule(above, equal_average, std_dev)


def duplicate_values(*, unique: bool = False) -> Any:
    return XLDuplicateValuesRule(unique)


def contains_text(text: str) -> Any:
    return XLContainsTextRule(text)


def not_contains_text(text: str) -> Any:
    return XLNotContainsTextRule(text)


def contains_blanks() -> Any:
    return XLContainsBlanksRule()


def not_contains_blanks() -> Any:
    return XLNotContainsBlanksRule()


def contains_errors() -> Any:
    return XLContainsErrorsRule()


def not_contains_errors() -> Any:
    return XLNotContainsErrorsRule()


_OP_ALIASES = {
    "<": "LessThan",
    "<=": "LessThanOrEqual",
    "=": "Equal",
    "==": "Equal",
    "!=": "NotEqual",
    "<>": "NotEqual",
    ">=": "GreaterThanOrEqual",
    ">": "GreaterThan",
    "between": "Between",
    "notbetween": "NotBetween",
    "not_between": "NotBetween",
}


def _parse_cf_operator(operator: Union[str, XLCfOperator]) -> XLCfOperator:
    if isinstance(operator, XLCfOperator):
        return operator
    key = operator.strip()
    name = _OP_ALIASES.get(key, _OP_ALIASES.get(key.lower(), key))
    # XLCfOperator member lookup
    if hasattr(XLCfOperator, name):
        return getattr(XLCfOperator, name)
    # Try case-insensitive enum name
    for attr in dir(XLCfOperator):
        if attr.lower() == name.lower() and not attr.startswith("_"):
            return getattr(XLCfOperator, attr)
    raise ValueError(f"Unknown CF operator: {operator!r}")
