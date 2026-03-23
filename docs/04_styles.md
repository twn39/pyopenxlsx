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
