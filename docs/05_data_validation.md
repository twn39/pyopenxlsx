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

---

## Properties and Methods of `DataValidation`

Once a validation rule is created, you can interact with its underlying C++ properties directly:
- **Properties/Methods**: `sqref`, `type`, `operator`, `allow_blank`, `show_drop_down`, `show_input_message`, `show_error_message`, `ime_mode`, `formula1`, `formula2`.
- **String Getters**: `prompt_title`, `prompt`, `error_title`, `error`, `error_style`.
- **Helper Methods**: `add_cell(ref)`, `add_range(ref)`, `set_list(["a", "b"])`, `set_reference_drop_list("Sheet", "A1:A3")`, `set_prompt(title, msg)`, `set_error(title, msg, style)`.

## `DataValidations` Collection

The `ws.data_validations` property provides a list-like collection.
- **`append()`**: Adds an empty validation and returns it.
- **`add_validation(...)`**: Quick-adds a populated validation.
- **`remove(index_or_sqref)`**: Deletes a specific validation.
- **`clear()`**: Deletes all validations on the sheet.

## Advanced Example: Dependent Dropdowns (Cascading Validation)
*Note: Excel evaluates formulas in data validation contextually. To do dependent dropdowns, we use the `INDIRECT` function pointing to another cell.*

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    
    # Setup lookup lists using Defined Names
    # 'Fruits' list
    ws["X1"].value = "Apple"
    ws["X2"].value = "Banana"
    wb.defined_names.append("Fruits", "Sheet1!$X$1:$X$2")
    
    # 'Cars' list
    ws["Y1"].value = "Toyota"
    ws["Y2"].value = "Ford"
    wb.defined_names.append("Cars", "Sheet1!$Y$1:$Y$2")
    
    # 1. Primary Dropdown in A1
    ws.data_validations.add_validation(
        "A1", 
        type="list", 
        formula1='"Fruits,Cars"', # The literal string list
        show_drop_down=True
    )
    
    # 2. Dependent Dropdown in B1
    # Reads the text in A1, and uses INDIRECT to find the defined name
    ws.data_validations.add_validation(
        "B1",
        type="list",
        formula1="INDIRECT($A$1)", 
        show_drop_down=True
    )
    
    wb.save("dependent_dropdown.xlsx")
```
