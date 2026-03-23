# Rich Text & Comments API

## Rich Text

Rich Text allows a single cell to contain multiple text segments, each with its own font formatting (e.g., mixing bold, italic, and different colors within one cell).

```python
from pyopenxlsx import Workbook, XLRichText, XLRichTextRun, XLColor

wb = Workbook()
ws = wb.active

# 1. Create a Rich Text object
rt = XLRichText()

# 2. Add normal text run
run1 = XLRichTextRun("Normal text, ")
rt.add_run(run1)

# 3. Add styled text runs
run2 = XLRichTextRun("Bold Red")
run2.bold = True
run2.font_color = XLColor(255, 0, 0)
rt.add_run(run2)

run3 = XLRichTextRun(", and ")
rt.add_run(run3)

run4 = XLRichTextRun("Italic Blue.")
run4.italic = True
run4.font_color = XLColor(0, 0, 255)
rt.add_run(run4)

# 4. Assign the XLRichText object directly to the cell's value
ws["A1"].value = rt

wb.save("richtext.xlsx")
```

## Comments

Cell comments are notes attached to individual cells. `pyopenxlsx` automatically resizes comment boxes to fit the text perfectly.

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# Single line comment
ws["A1"].value = "Look here"
ws["A1"].comment = "This is an auto-sized comment."

# Multiline comment
ws["B2"].comment = "First line of comment.\nSecond line of comment.\nPerfectly sized!"

# Remove a comment
ws["A1"].comment = None

wb.save("comments.xlsx")
```
