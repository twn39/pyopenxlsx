# Conditional Formatting API

`pyopenxlsx` supports conditional formatting rules (color scales, data bars, cell comparisons, formulas, icon sets, and more). Prefer builders in `pyopenxlsx.conditional_formatting` for concise call sites; native `XL*` factories remain available.

## High-level builders (recommended)

```python
from pyopenxlsx import Workbook, conditional_formatting as cf

with Workbook() as wb:
    ws = wb.active
    ws.write_rows(1, [[1, 2, 3], [4, 5, 6], [7, 8, 9]])

    # Hex or RGB tuples
    ws.add_conditional_formatting("A1:C1", cf.color_scale("#FF0000", "#00FF00"))
    ws.add_conditional_formatting("A2:C2", cf.data_bar((0, 0, 255), show_value=True))
    ws.add_conditional_formatting("A3:C3", cf.cell_is(">", "5"))
    ws.add_conditional_formatting("A1:C3", cf.formula_rule("A1>5"))
    ws.add_conditional_formatting("A1:C3", cf.top10(3))
    ws.add_conditional_formatting("A1:C3", cf.icon_set("3TrafficLights1"))

    wb.save("conditional_formatting.xlsx")
```

### Builder functions

| Function | Description |
| --- | --- |
| `color_scale(start, end, mid=None)` | 2- or 3-stop colour scale |
| `data_bar(color, *, show_value=True)` | Data bars |
| `cell_is(operator, formula, formula2=None)` | Comparison (`">"`, `">="`, `"between"`, … or `XLCfOperator`) |
| `formula_rule(formula)` | Expression rule |
| `icon_set` / `top10` / `above_average` | Rank and icon rules |
| `contains_text` / `duplicate_values` / blanks / errors | Text and uniqueness rules |

## Native rules

```python
from pyopenxlsx._openxlsx import XLColorScaleRule, XLDataBarRule, XLColor

scale = XLColorScaleRule(XLColor(255, 0, 0), XLColor(0, 255, 0))
ws.add_conditional_formatting("A1:C1", scale)
bar = XLDataBarRule(XLColor(0, 0, 255), show_value=True)
ws.add_conditional_formatting("A2:C2", bar)
```

## Managing rules

- `ws.add_conditional_formatting(sqref, rule)`
- `ws.remove_conditional_formatting(sqref)`
- `ws.clear_all_conditional_formatting()`
