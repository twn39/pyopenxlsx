# Pivot Tables API

`pyopenxlsx` provides bindings to generate Data Pivot Tables directly from source data. It also supports adding **Slicers** to dynamically filter pivot tables or standard tables.

> **Important Setup Rule:** When generating a pivot table from scratch using this API, it is highly recommended to place the pivot table on a **different worksheet** from the source data, and you must ensure the `target_cell` does NOT contain the worksheet name.

## Code Example

With the modern `XLPivotTableOptions` fluent builder API, configuring complex pivot tables is straightforward and memory-safe:

```python
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLPivotTableOptions, XLPivotSubtotal, XLSlicerOptions

with Workbook() as wb:
    # 1. Write source data to a sheet
    ws_data = wb.active
    ws_data.name = "SalesData"
    ws_data.write_row(1, ["Date", "Region", "Product", "Sales Rep", "Units", "Revenue"])
    ws_data.write_rows(2, [
        ["2024-01-01", "North", "Laptop", "Alice", 50, 50000.0],
        ["2024-01-02", "South", "Laptop", "Alice", 20, 20000.0],
        ["2024-01-03", "North", "Mouse",  "Bob",   300, 6000.0]
    ])
    
    # 2. Create a separate sheet for the Pivot Table
    ws_pivot = wb.create_sheet("PivotSheet")
    
    # 3. Configure options using the Fluent Builder API
    # The source must include the sheet name!
    # The target must ONLY be the cell coordinate (no sheet name!)
    options = XLPivotTableOptions("SalesPivot", "SalesData!A1:F4", "B3")
    
    # Chain methods to build the pivot table configuration
    (options
        .add_filter_field("Date")
        .add_row_field("Region")
        .add_row_field("Sales Rep")
        .add_column_field("Product")
        .add_data_field("Units", "Total Units Sold", XLPivotSubtotal.Sum, 3)     # 3 = '#,##0' format
        .add_data_field("Revenue", "Total Revenue ($)", XLPivotSubtotal.Sum, 4)  # 4 = '#,##0.00' format
        .set_pivot_table_style("PivotStyleMedium14")
        .set_show_row_stripes(True)
        .set_compact_data(True)
    )
    
    # 4. Add the pivot table to the new sheet
    ws_pivot._sheet.add_pivot_table(options)

    # 5. Add a Pivot Slicer for the "Region" column
    slicer_opts = XLSlicerOptions()
    slicer_opts.name = "RegionSlicer"
    slicer_opts.caption = "Filter by Region"
    
    # Access the newly created pivot table object by name
    pivot_table = ws_pivot._sheet.get_pivot_table("SalesPivot")
    ws_pivot._sheet.add_pivot_slicer("E3", pivot_table, "Region", slicer_opts)
    
    wb.save("pivot_demo.xlsx")
```

## `XLPivotTableOptions` Fluent Methods

The `XLPivotTableOptions` object provides various chainable methods:

- `add_row_field(field_name: str)`: Adds a field to the rows area.
- `add_column_field(field_name: str)`: Adds a field to the columns area.
- `add_filter_field(field_name: str)`: Adds a field to the report filter area.
- `add_data_field(field_name: str, custom_name: str, subtotal: XLPivotSubtotal, num_fmt_id: int)`: Adds a field to the values area.
  - `subtotal`: The aggregation type (e.g., `XLPivotSubtotal.Sum`, `XLPivotSubtotal.Count`, `XLPivotSubtotal.Average`).
  - `custom_name`: Overrides the default "Sum of X" text in the UI.
  - `num_fmt_id`: Standard Excel number format ID.

### Layout & Style Methods
- `set_pivot_table_style(style_name: str)`: Apply a predefined Excel style (e.g., `"PivotStyleLight16"`).
- `set_show_row_stripes(value: bool)`: Enable/disable banded rows.
- `set_show_col_stripes(value: bool)`: Enable/disable banded columns.
- `set_row_grand_totals(value: bool)`: Show/hide grand totals for rows.
- `set_col_grand_totals(value: bool)`: Show/hide grand totals for columns.
- `set_compact_data(value: bool)`: Enable/disable compact layout.

## `XLSlicerOptions` Configuration
- `name`: Internal unique name for the slicer cache.
- `caption`: The visible title displayed on the slicer UI.
- `width` / `height`: Dimensions of the slicer bounding box.
- `offset_x` / `offset_y`: Fine-tune the position relative to the target cell.
