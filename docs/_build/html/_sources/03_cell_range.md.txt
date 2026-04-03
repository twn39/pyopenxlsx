# Cell, Range, and Column API

## `Cell`

Represents a single Excel cell.

### Properties

- **`value`** (`Any`): The data stored in the cell (string, int, float, boolean, or datetime).
- **`formula`** (`Formula`): The formula object for the cell. Read/write string representations without the leading `=`.
- **`style_index`** / **`style`** (`int`): The integer ID of the style applied to the cell.
- **`is_date`** (`bool`): Returns `True` if the cell's number format indicates it is a date/time.
- **`comment`** (`str`): Get or set the text of a comment attached to the cell.

### Read-Only Style Properties
- **`font`**: Underlying `XLFont` object.
- **`fill`**: Underlying `XLFill` object.
- **`border`**: Underlying `XLBorder` object.
- **`alignment`**: Underlying `XLAlignment` object.

### Example
```python
cell = ws.cell(1, 1)
cell.value = 100
cell.style_index = my_style_id
cell.comment = "This is a comment"
```

---

## `Range`

Represents a block of cells. Obtained via `ws.range("A1:C3")`.

### Properties
- **`address`** (`str`): The string reference of the range (e.g., `"A1:C3"`).
- **`num_rows`** (`int`): Number of rows in the range.
- **`num_columns`** (`int`): Number of columns in the range.

### Methods
- **`clear()`**: Clears data and formulas from all cells in the range.
- **`__iter__()`**: Yields each `Cell` in the range, scanning left-to-right, top-to-bottom.

---

## `Column`

Used to adjust column-specific properties. Obtained via `ws.column("A")` or `ws.column(1)`.

### Properties
- **`width`** (`float`): The width of the column.
- **`hidden`** (`bool`): Whether the column is hidden.
- **`style_index`** (`int`): The default style for the column.

## Advanced Example: Formulas and Merging
```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    
    # 1. Writing data
    ws["A1"].value = "Revenue"
    ws["B1"].value = 1000
    ws["B2"].value = 2500
    ws["B3"].value = 1500
    
    # 2. Applying a Formula
    # IMPORTANT: Do not include the leading '=' in the formula string
    ws["B4"].formula = "SUM(B1:B3)"
    
    # 3. Using Ranges to clear or format blocks
    # Clear the text we just wrote in B3
    ws.range("B3:B3").clear() 
    
    # 4. Merging cells for a title header
    ws.merge_cells("C1:E1")
    ws["C1"].value = "Q1 Highlights"
    
    # Check if a cell is part of a merge
    if "C1:E1" in ws.merges:
        print("Range is merged successfully!")
        
    wb.save("cell_ops.xlsx")
```
