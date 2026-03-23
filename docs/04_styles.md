# Styles API

`pyopenxlsx` provides a declarative style API. You define style components, combine them via `Workbook.add_style()`, and apply the resulting integer `style_index` to cells.

## Components

### `Font`
```python
from pyopenxlsx import Font
font = Font(
    name="Calibri", 
    size=12, 
    bold=True, 
    italic=False, 
    underline="single", 
    color="FF0000" # AARRGGBB or RRGGBB hex
)
```

### `Fill`
```python
from pyopenxlsx import Fill
fill = Fill(pattern_type="solid", color="FFFF00") # Yellow background
```

### `Border` & `Side`
```python
from pyopenxlsx import Border, Side
side = Side(style="thin", color="000000")
border = Border(left=side, right=side, top=side, bottom=side)
```

### `Alignment`
```python
from pyopenxlsx import Alignment
align = Alignment(horizontal="center", vertical="center", wrap_text=True)
```

### `Protection`
```python
from pyopenxlsx import Protection
prot = Protection(locked=True, hidden=True)
```

---

## Applying Styles

1. Combine the components using `Workbook.add_style()`.
2. Assign the returned index to a `Cell`, `Row`, or `Column`.

```python
from pyopenxlsx import Workbook, Font, Fill

wb = Workbook()
ws = wb.active

# Register style
my_style_id = wb.add_style(
    font=Font(bold=True), 
    fill=Fill(pattern_type="solid", color="EEEEEE"),
    number_format="0.00%" # Built-in or custom format string
)

# Apply to a cell
ws["A1"].value = 0.85
ws["A1"].style_index = my_style_id
```

---

## Advanced Properties & Getters/Setters
*Note: Under the hood, these classes act as proxies to C++ OpenXLSX objects. While it's best to initialize them via `__init__`, they also expose native getter/setter methods (e.g., `set_name()`, `set_size()`, `set_color()`, `set_pattern_type()`, `set_bold()`, `set_italic()`, `set_left()`, `set_right()`, `set_top()`, `set_bottom()`, `set_horizontal()`, `set_vertical()`, `set_wrap_text()`) for fine-grained manipulation.*

## Advanced Example: Complex Financial Styling
```python
from pyopenxlsx import Workbook, Style, Font, Fill, Border, Side, Alignment

with Workbook() as wb:
    ws = wb.active
    
    # 1. Create a "Header" style
    # Deep blue background, white bold text, centered, with bottom border
    header_style_id = wb.add_style(
        font=Font(color="FFFFFF", bold=True, size=12),
        fill=Fill(pattern_type="solid", color="1F497D"),
        alignment=Alignment(horizontal="center", vertical="center"),
        border=Border(bottom=Side(style="medium", color="000000"))
    )
    
    # 2. Create a "Currency" style
    # Right-aligned, formatted as currency, thin borders everywhere
    thin_border = Side(style="thin", color="D9D9D9")
    currency_style_id = wb.add_style(
        font=Font(name="Consolas", size=11),
        alignment=Alignment(horizontal="right"),
        border=Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border),
        number_format='"$"#,##0.00' # Custom Excel format string
    )
    
    # Apply to row 1
    ws.write_row(1, ["Quarter", "Revenue", "Profit"])
    for col in range(1, 4):
        ws.cell(1, col).style_index = header_style_id
        
    # Apply to data
    ws.write_row(2, ["Q1", 50000.5, 12500])
    ws.cell(2, 2).style_index = currency_style_id
    ws.cell(2, 3).style_index = currency_style_id
    
    wb.save("styled_financials.xlsx")
```
