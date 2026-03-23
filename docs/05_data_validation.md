# Data Validation API

Data validation restricts what users can input into specific cells (e.g., dropdowns, numeric constraints).

## Core Concepts

Access a worksheet's validations via `ws.data_validations`.

### Adding Validations: Quick Method

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# Dropdown list
ws.data_validations.add_validation(
    "A1:A10", 
    type="list", 
    formula1='"Item 1,Item 2,Item 3"', 
    show_drop_down=True
)

# Numeric constraint
ws.data_validations.add_validation(
    "B1:B10",
    type="whole",
    operator="between",
    formula1="1",
    formula2="100"
)
```

### Adding Validations: Detailed Method

```python
from pyopenxlsx import XLDataValidationType, XLDataValidationOperator

dv = ws.data_validations.append()
dv.sqref = "C1:C10"
dv.type = XLDataValidationType.Decimal
dv.operator = XLDataValidationOperator.GreaterThan
dv.formula1 = "0.0"

# Custom Messages
dv.set_prompt("Input Required", "Enter a positive number.")
dv.set_error("Invalid Input", "Value must be strictly greater than 0.", style="stop")
```

## Enums

### `XLDataValidationType`
- `None_`, `Custom`, `Date`, `Decimal`, `List`, `TextLength`, `Time`, `Whole`

### `XLDataValidationOperator`
- `Between`, `NotBetween`, `Equal`, `NotEqual`, `GreaterThan`, `LessThan`, `GreaterThanOrEqual`, `LessThanOrEqual`
