# Page Setup & Printing API

`pyopenxlsx` allows you to configure how a worksheet is printed or exported to PDF using `page_setup`, `page_margins`, and `print_options`.

## Example Configuration

```python
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLPageOrientation

with Workbook() as wb:
    ws = wb.active
    
    # -----------------------------
    # 1. Page Setup
    # -----------------------------
    setup = ws.page_setup
    setup.orientation = XLPageOrientation.Landscape
    setup.paper_size = 9 # 9 = A4
    setup.scale = 80     # Print at 80% scale
    setup.black_and_white = True
    
    # -----------------------------
    # 2. Page Margins (in inches)
    # -----------------------------
    margins = ws.page_margins
    margins.left = 0.5
    margins.right = 0.5
    margins.top = 0.75
    margins.bottom = 0.75
    margins.header = 0.3
    margins.footer = 0.3
    
    # -----------------------------
    # 3. Print Options
    # -----------------------------
    options = ws.print_options
    options.grid_lines = True         # Print gridlines
    options.headings = True           # Print row/column headers (A, B, 1, 2)
    options.horizontal_centered = True
    options.vertical_centered = False
    
    # -----------------------------
    # 4. Print Area & Titles
    # -----------------------------
    ws.set_print_area("A1:E50")
    
    # Repeat specific rows/cols on every printed page
    ws.set_print_title_rows(1, 2) # Repeat rows 1-2
    ws.set_print_title_cols(1, 1) # Repeat column A (1)
    
    wb.save("printing_setup.xlsx")
```
