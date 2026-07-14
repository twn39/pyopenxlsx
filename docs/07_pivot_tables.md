# Pivot Tables API

`pyopenxlsx` can create Data Pivot Tables from source ranges and attach **Slicers**. Prefer the high-level `PivotTableBuilder` for everyday use; native `XLPivotTableOptions` remains available for advanced control.

> **Important Setup Rule:** When generating a pivot table from scratch, place it on a **different worksheet** from the source data, and ensure `target_cell` does **not** include a worksheet name (cell only, e.g. `"B3"`). The **source** range **must** include the sheet name (e.g. `"SalesData!A1:F4"`).

## High-level builder (recommended)

```python
from pyopenxlsx import Workbook, PivotTableBuilder, pivot_table

with Workbook() as wb:
    ws_data = wb.active
    ws_data.name = "SalesData"
    ws_data.write_row(1, ["Date", "Region", "Product", "Sales Rep", "Units", "Revenue"])
    ws_data.write_rows(2, [
        ["2024-01-01", "North", "Laptop", "Alice", 50, 50000.0],
        ["2024-01-02", "South", "Laptop", "Alice", 20, 20000.0],
        ["2024-01-03", "North", "Mouse",  "Bob",   300, 6000.0],
    ])

    ws_pivot = wb.create_sheet("PivotSheet")

    (
        PivotTableBuilder("SalesPivot", "SalesData!A1:F4", "B3")
        .filters("Date")
        .rows("Region", "Sales Rep")
        .columns("Product")
        .data("Units", name="Total Units Sold", subtotal="sum", num_fmt_id=3)
        .data("Revenue", name="Total Revenue ($)", subtotal="sum", num_fmt_id=4)
        .style("PivotStyleMedium14")
        .stripes(rows=True)
        .compact(True)
        .add_to(ws_pivot)
    )

    # One-shot helper
    # pivot_table("P2", "SalesData!A1:F4", "H3", rows="Region", data="Revenue").add_to(ws_pivot)

    wb.save("pivot_demo.xlsx")
```

`ws.add_pivot_table(...)` accepts either a `PivotTableBuilder` or native `XLPivotTableOptions`.

### `PivotTableBuilder` methods

| Method | Purpose |
| --- | --- |
| `rows(*fields)` / `columns(*fields)` / `filters(*fields)` | Layout fields |
| `data(field, *, name="", subtotal="sum", num_fmt_id=0)` | Values area (`subtotal` accepts `"sum"`, `"count"`, … or `XLPivotSubtotal`) |
| `style(name)` | Excel pivot style name |
| `stripes` / `show_headers` / `grand_totals` / `compact` / `data_on_rows` | Layout flags |
| `configure(**flags)` | Pass-through to native `set_*` booleans |
| `add_to(worksheet)` | Create the pivot on the sheet |

## Native `XLPivotTableOptions` (advanced)

```python
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLPivotTableOptions, XLPivotSubtotal, XLSlicerOptions

with Workbook() as wb:
    ws_data = wb.active
    ws_data.name = "SalesData"
    # ... write source data ...
    ws_pivot = wb.create_sheet("PivotSheet")

    options = XLPivotTableOptions("SalesPivot", "SalesData!A1:F4", "B3")
    (
        options
        .add_filter_field("Date")
        .add_row_field("Region")
        .add_column_field("Product")
        .add_data_field("Revenue", "Total Revenue ($)", XLPivotSubtotal.Sum, 4)
        .set_pivot_table_style("PivotStyleMedium14")
    )
    ws_pivot.add_pivot_table(options)

    slicer_opts = XLSlicerOptions()
    slicer_opts.name = "RegionSlicer"
    slicer_opts.caption = "Filter by Region"
    pivot = ws_pivot._sheet.get_pivot_table("SalesPivot")
    ws_pivot._sheet.add_pivot_slicer("E3", pivot, "Region", slicer_opts)

    wb.save("pivot_native.xlsx")
```

### Native fluent methods

- `add_row_field` / `add_column_field` / `add_filter_field` / `add_data_field`
- Layout: `set_pivot_table_style`, `set_show_row_stripes`, `set_compact_data`, grand totals, etc.

### `XLSlicerOptions`

- `name`, `caption`, `width` / `height`, `offset_x` / `offset_y`
