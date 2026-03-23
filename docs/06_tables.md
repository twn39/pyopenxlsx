# Tables (ListObjects) API

Excel Tables (ListObjects) allow data to be managed independently of the rest of the worksheet, complete with automatic filtering, sorting, and styling (e.g., row stripes).

## Creating a Table

To create a table, you access `ws.table` or `ws.add_table()`.

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# 1. Write the headers and data first
headers = ["ID", "Name", "Score"]
ws.write_row(1, headers)
ws.write_rows(2, [
    [1, "Alice", 95],
    [2, "Bob", 82],
    [3, "Charlie", 88]
])

# 2. Define the table
table = ws.add_table("StudentTable", "A1:C4")

# 3. Apply Table Styling
table.style_name = "TableStyleMedium2"
table.show_row_stripes = True

# 4. (Optional) Columns are auto-populated from worksheet headers when add_table is called.

wb.save("tables.xlsx")
```

## Table Properties

- **`name`** (`str`): The internal name of the table.
- **`display_name`** (`str`): The name shown in the Excel UI.
- **`range_reference`** (`str`): The address of the table (e.g., `"A1:C4"`).
- **`style_name`** (`str`): The built-in Excel style (e.g., `"TableStyleLight1"`).
- **`show_row_stripes`** (`bool`): Alternating row colors.
- **`show_column_stripes`** (`bool`): Alternating column colors.
- **`show_first_column`** / **`show_last_column`** (`bool`): Emphasize the first or last column.
- **`show_totals_row`** (`bool`): Display a totals row at the bottom.

## Table Columns

When `add_table()` is called, OpenXLSX automatically reads the first row of your specified range and creates columns matching those text values. You usually do not need to call `append_column` manually unless you are building a table structure programmatically before writing data to the worksheet.

### `append_column(name: str)`
Manually appends a new column to the table definition.

## Advanced Table Properties

- **`range`**: Alias for `range_reference`.
- **`style`**: Alias for `style_name`.
