# Charts API

`pyopenxlsx` provides basic support for inserting various types of charts into worksheets. 
*Note: The Chart API is currently exposed directly through the underlying C++ `_sheet` object and involves specifying raw chart types.*

## Adding a Chart

```python
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLChartType

with Workbook() as wb:
    ws = wb.active
    
    # 1. Write the chart's data source
    ws.write_rows(1, [
        ["Category", "Series 1", "Series 2"],
        ["A", 10, 20],
        ["B", 15, 25],
        ["C", 20, 30]
    ])
    
    # 2. Add the chart
    # add_chart(type, name, row, col, width, height)
    # The row and col dictate the top-left anchor of the chart.
    chart = ws._sheet.add_chart(
        XLChartType.Bar, # Type
        "MyChart",       # Internal Name
        5,               # Row anchor
        5,               # Column anchor
        400,             # Width in pixels
        300              # Height in pixels
    )
    
    wb.save("chart.xlsx")
```

## Supported Chart Types (`XLChartType`)
- `Bar`, `BarStacked`, `BarPercentStacked`, `Bar3D`
- `Line`, `LineStacked`, `LinePercentStacked`, `Line3D`
- `Pie`, `Pie3D`, `Doughnut`
- `Scatter`
- `Area`, `AreaStacked`, `AreaPercentStacked`, `Area3D`
- `Radar`, `RadarFilled`, `RadarMarkers`

*(Note: Data mapping to charts relies on the current defaults of the OpenXLSX backend; more granular series configuration will be available in future wrappers).*
