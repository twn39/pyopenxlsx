# Worksheet API

The `Worksheet` class represents a single tab within a `Workbook`. It is used to manipulate cells, rows, columns, and sheet-level configurations.

## Properties

- **`title`** (`str`): The name of the worksheet.
- **`index`** (`int`): The 0-based position of the sheet in the workbook.
- **`sheet_state`** (`str`): Visibility state (`"visible"`, `"hidden"`, `"very_hidden"`).
- **`max_row`** (`int`): The highest row number containing data.
- **`max_column`** (`int`): The highest column number containing data.
- **`merges`** (`MergeCells`): Collection of merged cell ranges.
- **`auto_filter`** (`str`): Get or set the autofilter range (e.g., `"A1:D10"`).
- **`zoom`** (`int`): The zoom scale percentage (e.g., `150` for 150%).
- **`has_panes`** (`bool`): Whether the sheet has frozen or split panes.
- **`page_setup`**, **`page_margins`**, **`print_options`**: Objects to configure printing behavior.

---

## Data Access & Modification Methods

### `cell(row: int, column: int, value: Any = None) -> Cell`
Retrieves a cell by its 1-based row and column index. Optionally sets its value.
```python
ws.cell(row=1, column=1, value="Header")
```

### `range(address: str) -> Range`
Retrieves a range of cells.
```python
rng = ws.range("A1:C3")
for cell in rng:
    cell.value = 0
```

### `append(iterable)`
Appends a single row of data to the bottom of the current sheet.
```python
ws.append(["Name", "Age", "City"])
```

### `write_row(row: int, values: list, start_col: int = 1)`
Writes a list of values to a specific row.

### `write_rows(start_row: int, data: list[list], start_col: int = 1)`
Writes a 2D list of data starting at a specific cell. Highly optimized for speed.

### `set_cell_value(row: int, col: int, value: Any)`
Directly sets a cell's value bypassing Python object creation. (Maximum performance).

### `get_cell_value(row: int, col: int) -> Any`
Directly gets a cell's value.

### `get_rows_data() -> list[list[Any]]`
Extracts all data from the sheet into a 2D Python list.

---

## Formatting & View Methods

### `set_column_format(col: int | str, style_idx: int)`
Sets the default style for an entire column.

### `set_row_format(row: int, style_idx: int)`
Sets the default style for an entire row.

### `insert_row(row_number: int, count: int = 1)`
Inserts one or more rows at the given row number (1-based index).
```python
ws.insert_row(5, count=2)  # Inserts 2 empty rows at row 5
```

### `delete_row(row_number: int, count: int = 1)`
Deletes one or more rows starting at the given row number.

### `insert_column(col_number: int, count: int = 1)`
Inserts one or more columns at the given column number (1-based index).

### `delete_column(col_number: int, count: int = 1)`
Deletes one or more columns starting at the given column number.

### `freeze_panes(ref_or_row, col=None)`
Freezes the view. `ws.freeze_panes("B2")` freezes row 1 and column A.

### `split_panes(x_split: float, y_split: float, top_left_cell: str, active_pane)`
Splits the view into scrollable panes.

### `add_image(img_path: str, anchor: str, width=None, height=None)`
Inserts an image into the worksheet.

### `add_shape(row: int, col: int, shape_type: str, **kwargs)`
Adds a vector shape to the worksheet. Supports over 20 shape types (e.g. `"Rectangle"`, `"Arrow"`, `"Cloud"`, `"FlowChartProcess"`) and configurable appearance options (like `fill_color`, `line_color`, `rotation`, `text`).

### `protect(password: str, **options)`
Protects the worksheet. Supports highly granular protection options (e.g. `format_cells`, `insert_columns`, `sort`, `auto_filter`, `objects`, `scenarios`, `select_locked_cells`, `select_unlocked_cells`, `delete_rows`, etc.).

### `auto_fit_column(col: int)`
Automatically adjusts the width of the specified column to fit its contents.
```python
ws.auto_fit_column(1) # Auto-fit column A
```
*(You can also use `ws.column("A").autofit()`)*

### `apply_auto_filter()`
Applies the autofilter dropdowns to the range specified in `ws.auto_filter`.
```python
ws.auto_filter = "A1:C10"
ws.apply_auto_filter()
```

## Hyperlinks & Comments

### `add_hyperlink(ref: str, url: str, tooltip: str = "")`
Adds an external hyperlink to a cell.

### `add_internal_hyperlink(ref: str, location: str, tooltip: str = "")`
Adds an internal link to another sheet or cell (e.g. `"Sheet2!A1"`).

### `has_hyperlink(ref: str) -> bool` / `get_hyperlink(ref: str) -> str` / `remove_hyperlink(ref: str)`
Helpers to check, retrieve, or remove hyperlinks.

### `add_comment(cell_ref: str, text: str, author: str)`
Adds a comment to a specific cell.

---

## Ranges & Merging

### `merge_cells(address: str)`
Merges a range of cells (e.g., `"A1:C3"`).

### `unmerge_cells(address: str)`
Unmerges a previously merged range.

---

## Tables & Shapes

- **`table`** / **`tables`**: Properties to access the Worksheet's ListObjects.
- **`add_table(name: str, range: str)`**: Creates a new Table.
- **`has_drawing`** / **`drawing`**: Check or access the underlying drawing object.
- **`add_sparkline(location: str, data_range: str, type)`**: Inserts a sparkline into a cell.

---

## Advanced I/O

- **`get_row_values(row: int) -> list[Any]`**: Gets a single row's values.
- **`iter_row_values()`**: Iterator yielding rows one by one.
- **`get_range_data(r1, c1, r2, c2)`** / **`get_range_values(...)`**: Bulk reading.
- **`write_range(r1, c1, data)`**: Optimized writing for numpy arrays/buffers.
- **`set_cells(cells: list[tuple])`**: Batch updates using a list of `(row, col, value)` tuples.

---

## Documented in Other Modules
- For conditional formatting: `add_conditional_formatting`, `remove_conditional_formatting`, `clear_all_conditional_formatting`.
- For streams: `stream_writer`, `stream_reader`.

### Missing Methods & Properties (Internal/Proxy)

- **`column(col)`**: Used internally to fetch a specific column.
- **`unprotect()`**: Disables worksheet protection.
- **`set_print_area(sqref: str)`**, **`set_print_title_rows(start: int, end: int)`**, **`set_print_title_cols(start: int, end: int)`**: Configures print properties directly on the worksheet.
- **`clear_panes()`**: Clears any existing split or frozen panes.
- **`rows`**: Yields an iterator for all rows containing data.
- **`data_validations`**: Property returning a `DataValidations` collection.
- **`protection`**: Property returning the current worksheet protection status in a dictionary.

## Advanced Example: Data Parsing and High-Speed Bulk Write
```python
from pyopenxlsx import Workbook
import datetime

with Workbook() as wb:
    ws = wb.active
    ws.title = "HighSpeedData"
    
    # 1. Formatting columns before data arrives
    # Apply a date format (14 is the built-in mm-dd-yy format index)
    style_date = wb.add_style(number_format=14)
    ws.set_column_format("C", style_date)
    
    # 2. Bulk Write using write_rows (Ultra-Fast)
    # write_rows bypasses Python Cell object creation
    data = [
        ["ID", "Name", "Join Date", "Score"],
        [1, "Alice", datetime.date(2023, 1, 15), 95.5],
        [2, "Bob", datetime.date(2023, 2, 20), 88.0],
        [3, "Charlie", datetime.date(2023, 3, 5), 92.2]
    ]
    ws.write_rows(1, data) # Start at Row 1
    
    # 3. View manipulations
    ws.auto_filter = "A1:D4"
    ws.apply_auto_filter()
    ws.auto_fit_column(2) # Auto-fit the "Name" column
    
    # 4. Sheet Level Protections
    ws.protect("my_secret", format_cells=True, sort=True)
    
    wb.save("bulk_data.xlsx")
```
