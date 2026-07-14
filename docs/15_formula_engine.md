# Formula Engine API

`pyopenxlsx` exposes the OpenXLSX-NX formula engine so you can evaluate Excel formulas in Python without Microsoft Excel. Two façades are provided:

| Class | Role |
| --- | --- |
| `FormulaEngine` | Stateless evaluation of a single formula string |
| `CalculationEngine` | Sheet/workbook recalculation with dependency tracking |

## FormulaEngine

```python
from pyopenxlsx import Workbook, FormulaEngine

engine = FormulaEngine()

# Basic evaluation (optional leading '=')
assert engine.evaluate("1 + 2 * 3") == 7
assert engine.evaluate('="Hello " & "World"') == "Hello World"
assert engine.evaluate("=IF(10 > 5, TRUE, FALSE)") is True
```

### Contextual evaluation

Pass a high-level `Worksheet` so cell references resolve against live data:

```python
from pyopenxlsx import Workbook, FormulaEngine

wb = Workbook()
ws = wb.active
ws["A1"].value = 10
ws["A2"].value = 20
ws["B1"].value = 5

engine = FormulaEngine()
assert engine.evaluate("SUM(A1:A2) * B1", ws) == 150
# or keyword form
assert engine.evaluate("SUM(A1:A2)", worksheet=ws) == 30
```

### Relative context (`ROW` / `COLUMN`)

```python
# Current cell context for parameterless ROW()/COLUMN()
engine.evaluate("ROW()", current_row=5, current_col=2, current_sheet="Sheet1")
```

### Convenience methods

```python
engine.sum("A1:A10", worksheet=ws)
engine.average("B1:B10", worksheet=ws)
engine.evaluate_many(["SUM(A1:A3)", "AVERAGE(A1:A3)"], worksheet=ws)
```

### `evaluate` signature

```text
evaluate(
    formula: str,
    worksheet=None,
    *,
    session=None,
    reporter=None,
    current_row=None,
    current_col=None,
    current_sheet=None,
) -> Any
```

- **formula**: With or without a leading `=`.
- **worksheet**: High-level `Worksheet` or `None` for pure expressions.
- **Returns**: Python scalar (`int` / `float` / `str` / `bool`), or raises on failure.

## CalculationEngine

Recalculate formulas stored in cells with dependency tracking:

```python
from pyopenxlsx import Workbook, CalculationEngine

wb = Workbook()
ws = wb.active
ws["A1"].value = 10
ws["A2"].formula = "A1*2"

calc = CalculationEngine(ws)  # or Workbook
calc.rebuild()
n = calc.recalculate()
assert ws["A2"].value == 20  # after write-back options / cell update semantics
```

### Methods

- `rebuild()`, `recalculate()`, `recalculate_all()`
- `calc_cell_value(a1)`, `mark_dirty(a1)`, `set_input_value(a1, value)`
- Properties: `formula_count`, `dirty_count`

Optional native `XLCalculationOptions` control write-back and defined-name usage.
