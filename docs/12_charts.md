# Charts API

`pyopenxlsx` supports common Excel chart types via the worksheet façade. Prefer string chart types and the fluent `Chart` wrapper; set `wrap=False` when you need the raw native object.

## Adding a chart (recommended)

```python
from pyopenxlsx import Workbook, add_chart

with Workbook() as wb:
    ws = wb.active
    ws.write_rows(1, [
        ["Category", "Series 1", "Series 2"],
        ["A", 10, 20],
        ["B", 15, 25],
        ["C", 20, 30],
    ])

    # String type names: "bar", "column", "line", "pie", ...
    chart = (
        ws.add_chart("bar", "MyChart", row=5, col=5, width=400, height=300)
        .title("Sales")
        .legend("bottom")
        .series(
            "Sheet1!$B$2:$B$4",
            name="Series 1",
            categories_ref="Sheet1!$A$2:$A$4",
        )
        .series(
            "Sheet1!$C$2:$C$4",
            name="Series 2",
            categories_ref="Sheet1!$A$2:$A$4",
        )
        .data_labels(value=True)
    )

    # Module-level helper with first series in one call
    add_chart(
        ws,
        "line",
        "LineChart",
        row=20,
        col=5,
        title="Trend",
        series_ref="Sheet1!$B$2:$B$4",
        cats_ref="Sheet1!$A$2:$A$4",
        legend="right",
    )

    wb.save("chart.xlsx")
```

### `Chart` fluent methods

| Method | Purpose |
| --- | --- |
| `title` / `legend` / `style` | Title, legend position (`"right"`, `"bottom"`, …), style id |
| `series` / `series_many` | Values + optional categories from A1 refs |
| `bubble_series` | Bubble charts |
| `data_labels` / `data_table` | Labels and data table |
| `overlap` / `gap_width` / `hole_size` / `rotation` | Chart-type-specific layout |
| `x_axis()` / `y_axis()` | Native axis objects for fine control |
| `raw` | Underlying native `XLChart` |

Native methods such as `set_title` / `add_series_ref` still work via attribute forwarding.

## Chart types

Friendly names resolve case-insensitively to `XLChartType`, including:

- Bar / Column family (stacked, percent, 3D)
- Line, Pie, Doughnut, Scatter variants
- Area, Radar, Bubble, Stock, Surface

```python
from pyopenxlsx import chart_type
from pyopenxlsx._openxlsx import XLChartType

assert chart_type("column") == XLChartType.Column
```

## Native path

```python
from pyopenxlsx._openxlsx import XLChartType

native = ws.add_chart(XLChartType.Bar, "C1", 5, 5, 400, 300, wrap=False)
native.set_title("Native")
native.add_series_ref("Sheet1!$B$2:$B$4", "S1", "Sheet1!$A$2:$A$4")
```
