# Pivot Tables API

`pyopenxlsx` provides bindings to generate Data Pivot Tables directly from source data. 

> **Important Setup Rule:** When generating a pivot table from scratch using this API, it is highly recommended to place the pivot table on a **different worksheet** from the source data, and you must ensure the `target_cell` does NOT contain the worksheet name.

## Code Example

```python
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLPivotTableOptions, XLPivotField, XLPivotSubtotal

with Workbook() as wb:
    # 1. Write source data to a sheet
    ws_data = wb.active
    ws_data.title = "SalesData"
    ws_data.write_row(1, ["Region", "Product", "Sales"])
    ws_data.write_rows(2, [
        ["North", "Apples", 100],
        ["South", "Bananas", 300],
        ["North", "Oranges", 150]
    ])
    
    # 2. Create a separate sheet for the Pivot Table
    ws_pivot = wb.create_sheet("PivotSheet")
    
    # 3. Configure options
    options = XLPivotTableOptions()
    options.name = "SalesPivot"
    
    # The source must include the sheet name!
    options.source_range = "SalesData!A1:C4"
    
    # The target must ONLY be the cell coordinate (no sheet name!)
    options.target_cell = "A3" 
    
    # 4. Define fields
    # Row Field
    r = XLPivotField()
    r.name = "Region"
    r.subtotal = XLPivotSubtotal.Sum
    options.rows = [r]

    # Column Field
    c = XLPivotField()
    c.name = "Product"
    c.subtotal = XLPivotSubtotal.Sum
    options.columns = [c]

    # Data (Value) Field
    d = XLPivotField()
    d.name = "Sales"
    d.subtotal = XLPivotSubtotal.Sum
    d.custom_name = "Total Sales"
    options.data = [d]
    
    # 5. Add the pivot table to the new sheet
    ws_pivot._sheet.add_pivot_table(options)
    
    wb.save("pivot_demo.xlsx")
```

## `XLPivotField` Configuration
- `name`: Must exactly match the column header in the source data.
- `subtotal`: The aggregation type (e.g., `XLPivotSubtotal.Sum`, `XLPivotSubtotal.Count`, `XLPivotSubtotal.Average`).
- `custom_name`: Overrides the default "Sum of X" text in the UI (only applies to data fields).
