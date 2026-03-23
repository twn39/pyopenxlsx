# Conditional Formatting API

`pyopenxlsx` supports adding conditional formatting rules to highlight interesting cells, emphasize unusual values, and visualize data using data bars, color scales, and icon sets.

## Adding a Rule

You can apply conditional formatting to a worksheet range using `ws.add_conditional_formatting()`.

```python
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLColorScaleRule, XLDataBarRule, XLColor

with Workbook() as wb:
    ws = wb.active
    ws.write_rows(1, [[1, 2, 3], [4, 5, 6], [7, 8, 9]])
    
    # 1. Color Scale Rule (2-Color)
    # Highlights cells based on their value relative to the range
    color_min = XLColor(255, 0, 0) # Red
    color_max = XLColor(0, 255, 0) # Green
    scale_rule = XLColorScaleRule(color_min, color_max)
    ws.add_conditional_formatting("A1:C1", scale_rule)
    
    # 2. Data Bar Rule
    # Adds a horizontal bar inside the cell proportional to its value
    bar_color = XLColor(0, 0, 255) # Blue
    bar_rule = XLDataBarRule(bar_color, show_value_text=True)
    ws.add_conditional_formatting("A2:C2", bar_rule)
    
    wb.save("conditional_formatting.xlsx")
```

## Supported Rule Types

Currently, the underlying C++ engine exposes specific specialized rules via Python bindings:
- **`XLColorScaleRule(min_color: XLColor, max_color: XLColor)`**: Creates a gradient color scale between two colors.
- **`XLDataBarRule(color: XLColor, show_value: bool)`**: Creates a data bar with the specified color. If `show_value` is `False`, the underlying cell text is hidden.

*Note: More rule types (e.g., standard formula-based rules, icon sets) may be available through the underlying C++ API and will be fully exposed in future releases.*

## Managing Rules

- `ws.add_conditional_formatting(sqref: str, rule: XLCfRule)`: Applies a rule to the given range (e.g., `"A1:D10"`).
- `ws.remove_conditional_formatting(sqref: str)`: Removes all conditional formatting rules matching the exact range reference.
- `ws.clear_all_conditional_formatting()`: Clears all conditional formatting rules from the entire worksheet.
