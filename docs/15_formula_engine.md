# Formula Engine API

`pyopenxlsx` exposes the underlying OpenXLSX-NX C++ formula engine to allow static evaluation of Excel formulas within Python, without needing to open Microsoft Excel.

## FormulaEngine

The `FormulaEngine` class parses and evaluates formula strings.

```python
from pyopenxlsx import Workbook, FormulaEngine

engine = FormulaEngine()

# Basic evaluation (no cell references needed)
result = engine.evaluate("=1 + 2 * 3")
print(result) # Output: 7

# String operations
result = engine.evaluate('="Hello " & "World"')
print(result) # Output: "Hello World"

# Logical operations
result = engine.evaluate("=IF(10 > 5, TRUE, FALSE)")
print(result) # Output: True
```

### Contextual Evaluation (Resolving Cell References)

If your formula contains references to cells (like `A1`, `B2`), you must provide a `Worksheet` context so the engine can look up the values of those cells.

```python
from pyopenxlsx import Workbook, FormulaEngine

wb = Workbook()
ws = wb.active

# Populate some data
ws["A1"].value = 10
ws["A2"].value = 20
ws["B1"].value = 5

engine = FormulaEngine()

# Pass the worksheet context to evaluate
result = engine.evaluate("=SUM(A1:A2) * B1", worksheet=ws)
print(result) # Output: 150 ((10 + 20) * 5)
```

### Methods

#### `evaluate(formula: str, worksheet: Optional[Worksheet] = None) -> Any`
Evaluates the formula string.
- **Parameters:**
  - `formula`: The string to evaluate. Can start with or without the `=` sign.
  - `worksheet`: (Optional) The `pyopenxlsx.Worksheet` object to use for resolving cell references (e.g. `A1`).
- **Returns:** The calculated primitive Python value (e.g. `int`, `float`, `str`, `bool`), or raises an error if evaluation fails.
