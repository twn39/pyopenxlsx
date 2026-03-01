<div align="center">

# PyOpenXLSX

[![PyPI version](https://img.shields.io/pypi/v/pyopenxlsx.svg)](https://pypi.org/project/pyopenxlsx/)
[![Python versions](https://img.shields.io/pypi/pyversions/pyopenxlsx.svg)](https://pypi.org/project/pyopenxlsx/)
[![Downloads](https://img.shields.io/pypi/dm/pyopenxlsx.svg)](https://pypi.org/project/pyopenxlsx/)
[![Build Status](https://github.com/twn39/pyopenxlsx/actions/workflows/build.yml/badge.svg)](https://github.com/twn39/pyopenxlsx/actions/workflows/build.yml)
[![Codecov](https://codecov.io/gh/twn39/pyopenxlsx/branch/main/graph/badge.svg)](https://codecov.io/gh/twn39/pyopenxlsx)
[![License](https://img.shields.io/pypi/l/pyopenxlsx.svg)](https://opensource.org/licenses/BSD-3-Clause)

</div>

> [!IMPORTANT]
> `pyopenxlsx` currently uses a specialized **fork** of the [OpenXLSX](https://github.com/twn39/OpenXLSX) library (v1.0.0+), which includes critical performance optimizations and functional enhancements (such as custom properties and improved hyperlink handling) not currently available in the upstream repository.

`pyopenxlsx` is a high-performance Python binding for the [OpenXLSX](https://github.com/troldal/OpenXLSX) C++ library. It aims to provide significantly faster read/write speeds compared to pure Python libraries like `openpyxl`, while maintaining a Pythonic API design.

## Core Features

-   **High Performance**: Powered by the modern C++17 OpenXLSX library.
-   **Pythonic API**: Intuitive interface with properties, iterators, and context managers.
-   **Async Support**: `async/await` support for key I/O operations.
-   **Rich Styling**: Comprehensive support for fonts, fills, borders, alignments, and number formats.
-   **Extended Metadata**: Support for both standard and **custom document properties**.
-   **Advanced Content**: Support for **images**, **hyperlinks** (external/internal), and **comments**.
-   **Memory Safety**: Combines C++ efficiency with Python's automatic memory management.

## Tech Stack

| Component | Technology |
| :--- | :--- |
| **C++ Core** | [OpenXLSX](https://github.com/troldal/OpenXLSX) |
| **Bindings** | [nanobind](https://github.com/wjakob/nanobind) |
| **Build System** | [scikit-build-core](https://github.com/scikit-build/scikit-build-core) & [CMake](https://cmake.org/) |

## Installation

### From PyPI (Recommended)

```bash
# Using pip
pip install pyopenxlsx

# Using uv
uv pip install pyopenxlsx
```

### From Source

```bash
# Using uv
uv pip install .

# Or using pip
pip install .
```

### Development Installation

```bash
uv pip install -e .
```

## Quick Start

### Create and Save a Workbook

```python
from pyopenxlsx import Workbook

# Create a new workbook
with Workbook() as wb:
    ws = wb.active
    ws.title = "MySheet"
    
    # Write data
    ws["A1"].value = "Hello"
    ws["B1"].value = 42
    ws.cell(row=2, column=1).value = 3.14
    
    # Save
    wb.save("example.xlsx")
```

### Custom Properties

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    # Set custom document properties
    wb.custom_properties["Author"] = "Curry Tang"
    wb.custom_properties["Project"] = "PyOpenXLSX"
    wb.save("props.xlsx")
```

### Hyperlinks

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    ws["A1"].value = "Google"
    # External link
    ws.add_hyperlink("A1", "https://www.google.com", tooltip="Search")
    
    # Internal link to another sheet
    ws2 = wb.create_sheet("Data")
    ws["A2"].value = "See Data"
    ws.add_internal_hyperlink("A2", "Data!A1")
    
    wb.save("links.xlsx")
```

### Read a Workbook

```python
from pyopenxlsx import load_workbook

wb = load_workbook("example.xlsx")
ws = wb["MySheet"]
print(ws["A1"].value)  # Output: Hello
wb.close()
```

### Async Operations

`pyopenxlsx` provides `async/await` support for all I/O-intensive operations, ensuring your event loop remains responsive.

```python
import asyncio
from pyopenxlsx import Workbook, load_workbook_async, Font

async def main():
    # 1. Async context manager for automatic cleanup
    async with Workbook() as wb:
        ws = wb.active
        ws["A1"].value = "Async Data"
        
        # 2. Async stylesheet creation
        style_idx = await wb.add_style_async(font=Font(bold=True))
        ws["A1"].style_index = style_idx
        
        # 3. Async worksheet operations
        new_ws = await wb.create_sheet_async("AsyncSheet")
        await new_ws.append_async(["Dynamic", "Row", 123])
        
        # 4. Async range operations
        await new_ws.range("A1:C1").clear_async()
        
        # 5. Async save
        await wb.save_async("async_example.xlsx")

    # 6. Async load
    async with await load_workbook_async("async_example.xlsx") as wb:
        ws = wb.active
        print(ws["A1"].value)
        
        # 7. Async protection
        await ws.protect_async(password="secret")
        await ws.unprotect_async()

asyncio.run(main())
```

### Styling

```python
from pyopenxlsx import Workbook, Font, Fill, Border, Side, Alignment

wb = Workbook()
ws = wb.active

# Define styles using hex colors (ARGB) or names
# Hex colors can be 6-digit (RRGGBB) or 8-digit (AARRGGBB)
font = Font(name="Arial", size=14, bold=True, color="FF0000") # Red
fill = Fill(pattern_type="solid", color="FFFF00")              # Yellow
border = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin"),
    top=Side(style="thick"),
    bottom=Side(style="thin")
)
alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Apply style
style_idx = wb.add_style(font=font, fill=fill, border=border, alignment=alignment)
ws["A1"].value = "Styled Cell"
ws["A1"].style_index = style_idx

wb.save("styles.xlsx")
```

### Insert Images

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# Insert image at A1, automatically maintaining aspect ratio
# Requires Pillow: pip install pillow
ws.add_image("logo.png", anchor="A1", width=200)

# Or specify exact dimensions
ws.add_image("banner.jpg", anchor="B5", width=400, height=100)

wb.save("images.xlsx")
```

### Comments

Comments are **automatically resized** to fit their content by default.

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# Simple or multiline comments - all will auto-size perfectly
ws["A1"].comment = "Short comment"
ws["B2"].comment = "Line 1: High performance\nLine 2: Pythonic API\nLine 3: Auto-sized by default!"

wb.save("comments.xlsx")
```

---

## API Documentation

### Module Exports

```python
from pyopenxlsx import (
    # Core Classes
    Workbook, Worksheet, Cell, Range,
    load_workbook, load_workbook_async,
    
    # Style Classes
    Font, Fill, Border, Side, Alignment, Style, Protection,
    
    # Enums & Constants
    XLColor, XLSheetState, XLLineStyle, XLPatternType, XLAlignmentStyle,
    XLProperty, XLUnderlineStyle, XLFontSchemeStyle, XLVerticalAlignRunStyle,
    XLFillType,
)
```

---

### `Workbook` Class

The top-level container for an Excel file.

#### Constructor

```python
Workbook(filename: str | None = None, force_overwrite: bool = True)
```

-   `filename`: Optional. If provided, opens an existing file; otherwise, creates a new workbook.
-   `force_overwrite`: Optional (default: `True`). If `True`, overwrites existing files when creating or saving. If `False`, an exception is raised if the file already exists.

#### Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `active` | `Worksheet \| None` | Get or set the currently active worksheet. |
| `sheetnames` | `list[str]` | Returns a list of all worksheet names. |
| `properties` | `DocumentProperties` | Access standard document metadata (title, author, etc.). |
| `custom_properties` | `CustomProperties` | Access custom document properties (dict-like). |
| `styles` | `XLStyles` | Access underlying style object (advanced usage). |
| `workbook` | `XLWorkbook` | Access underlying C++ workbook object (advanced usage). |

#### Methods

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `save(filename=None, force_overwrite=True)` | `None` | Save the workbook. Saves to original path if `filename` is omitted. |
| `save_async(filename=None, force_overwrite=True)` | `Coroutine` | Asynchronously save the workbook. |
| `close()` | `None` | Close the workbook and release resources. |
| `close_async()` | `Coroutine` | Asynchronously close the workbook. |
| `create_sheet(title=None, index=None)` | `Worksheet` | Create a new worksheet. `title` defaults to "Sheet1", etc. |
| `create_sheet_async(...)` | `Coroutine` | Asynchronously create a worksheet. |
| `remove(worksheet)` | `None` | Delete the specified worksheet. |
| `remove_async(worksheet)` | `Coroutine` | Asynchronously delete a worksheet. |
| `copy_worksheet(from_worksheet)` | `Worksheet` | Copy a worksheet and return the copy. |
| `copy_worksheet_async(...)` | `Coroutine` | Asynchronously copy a worksheet. |
| `add_style(...)` | `int` | Create a new style and return its index. See below. |
| `add_style_async(...)` | `Coroutine` | Asynchronously create a style. |
| `get_embedded_images()` | `list[ImageInfo]` | Get list of all images embedded in the workbook. |
| `get_image_data(name)` | `bytes` | Get binary data of an embedded image by its name or path. |
| `extract_images(out_dir)` | `list[str]` | Extract all images to a directory. Returns list of file paths. |
| `extract_images_async(...)` | `Coroutine` | Asynchronously extract all images. |

#### `add_style` Method

```python
def add_style(
    font: Font | int | None = None,
    fill: Fill | int | None = None,
    border: Border | int | None = None,
    alignment: Alignment | None = None,
    number_format: str | int | None = None,
    protection: Protection | None = None,
) -> int:
```

**Returns:** A style index (`int`) that can be assigned to `Cell.style_index`.

**Example:**

```python
# Pass all styles via a Style object
from pyopenxlsx import Style, Font, Fill

style = Style(
    font=Font(bold=True),
    fill=Fill(color="C8C8C8"), # Hex color
    number_format="0.00"
)
idx = wb.add_style(style)
```

---

### `Worksheet` Class

Represents a sheet within an Excel file.

#### Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `title` | `str` | Get or set the worksheet name. |
| `index` | `int` | Get or set the worksheet index (0-based). |
| `sheet_state` | `str` | Visibility: `"visible"`, `"hidden"`, `"very_hidden"`. |
| `max_row` | `int` | Returns the maximum row index used. |
| `max_column` | `int` | Returns the maximum column index used. |
| `rows` | `Iterator` | Iterate over all rows with data. |
| `has_drawing` | `bool` | True if the worksheet has a drawing (images, etc.). |
| `drawing` | `XLDrawing` | Access underlying drawing object (advanced usage). |
| `merges` | `MergeCells` | Access merged cells information. |
| `protection` | `dict` | Get worksheet protection status (read-only). |

#### Methods

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `cell(row, column, value=None)` | `Cell` | Get cell by 1-based indices. Optionally set value. |
| `range(address)` | `Range` | Get range by string, e.g., `ws.range("A1:C3")`. |
| `range(start, end)` | `Range` | Get range by endpoints, e.g., `ws.range("A1", "C3")`. |
| `merge_cells(address)` | `None` | Merge cells, e.g., `ws.merge_cells("A1:B2")`. |
| `merge_cells_async(...)` | `Coroutine` | Asynchronously merge cells. |
| `unmerge_cells(address)` | `None` | Unmerge cells. |
| `unmerge_cells_async(...)` | `Coroutine` | Asynchronously unmerge cells. |
| `append(iterable)` | `None` | Append a row of data after the last used row. |
| `append_async(iterable)` | `Coroutine` | Asynchronously append a row. |
| `add_hyperlink(ref, url, tooltip="")` | `None` | Add an external hyperlink to a cell. |
| `add_internal_hyperlink(ref, loc, ...)` | `None` | Add an internal hyperlink (e.g., `"Sheet2!A1"`). |
| `set_column_format(col, style_idx)` | `None` | Set default style for a column. `col` can be int or "A". |
| `set_row_format(row, style_idx)` | `None` | Set default style for a row. |
| `column(col)` | `Column` | Get column object for width adjustments. |
| `protect(...)` | `None` | Protect the worksheet. |
| `protect_async(...)` | `Coroutine` | Asynchronously protect the worksheet. |
| `unprotect()` | `None` | Unprotect the worksheet. |
| `unprotect_async()` | `Coroutine` | Asynchronously unprotect. |
| `add_image(...)` | `None` | Insert an image. |
| `add_image_async(...)` | `Coroutine` | Asynchronously insert an image. |

#### `add_image` Method

```python
def add_image(
    img_path: str,
    anchor: str = "A1",
    width: int | None = None,
    height: int | None = None,
) -> None:
```

-   `img_path`: Path to image (PNG, JPG, GIF).
-   `anchor`: Top-left cell address.
-   `width`, `height`: Pixel dimensions. Requires Pillow for auto-detection if not provided.

#### Magic Methods

| Method | Description |
| :--- | :--- |
| `__getitem__(key)` | Get cell by address: `ws["A1"]` |

---

### `Cell` Class

The fundamental unit of data in Excel.

#### Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `value` | `Any` | Get/Set value. Supports `str`, `int`, `float`, `bool`, `datetime`. |
| `formula` | `Formula` | Get/Set formula string (without initial `=`). |
| `style_index` | `int` | Get/Set style index. |
| `style` | `int` | Alias for `style_index`. |
| `is_date` | `bool` | True if the cell has a date format. |
| `comment` | `str \| None` | Get/Set cell comment. Set `None` to remove. |
| `font` | `XLFont` | Get font object (read-only). |
| `fill` | `XLFill` | Get fill object (read-only). |
| `border` | `XLBorder` | Get border object (read-only). |
| `alignment` | `XLAlignment` | Get alignment object (read-only). |

#### Date Handling

If `is_date` is `True`, `value` automatically returns a Python `datetime` object.

```python
# Write
ws["A1"].value = datetime(2024, 1, 15)
ws["A1"].style_index = wb.add_style(number_format=14)  # Built-in date format

# Read
print(ws["A1"].value)   # datetime.datetime(2024, 1, 15, 0, 0)
print(ws["A1"].is_date) # True
```

#### Formulas

**Note**: Formulas must be set via the `formula` property, not `value`.

```python
# Correct
ws["A3"].formula = "SUM(A1:A2)" 

# Incorrect (treated as string)
ws["A3"].value = "=SUM(A1:A2)"
```

---

### `Range` Class

Represents a rectangular area of cells.

#### Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `address` | `str` | Range address, e.g., `"A1:C3"`. |
| `num_rows` | `int` | Row count. |
| `num_columns` | `int` | Column count. |

#### Methods

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `clear()` | `None` | Clear values in all cells of the range. |
| `clear_async()` | `Coroutine` | Asynchronously clear range. |

#### Iteration

```python
for cell in ws.range("A1:B2"):
    print(cell.value)
```

---

### Style Classes

#### `Font`
```python
Font(name="Arial", size=11, bold=False, italic=False, color=None)
```

#### `Fill`
```python
Fill(pattern_type="solid", color=None, background_color=None)
```

#### `Border`
```python
Border(left=Side(), right=Side(), top=Side(), bottom=Side(), diagonal=Side())
```

#### `Side`
```python
Side(style="thin", color=None)
```
**Styles**: `"thin"`, `"thick"`, `"dashed"`, `"dotted"`, `"double"`, `"hair"`, `"medium"`, `"mediumDashed"`, `"mediumDashDot"`, `"mediumDashDotDot"`, `"slantDashDot"`

#### `Alignment`
```python
Alignment(horizontal="center", vertical="center", wrap_text=True)
```
**Options**: `"left"`, `"center"`, `"right"`, `"general"`, `"top"`, `"bottom"`

---

### `DocumentProperties`

Accessed via `wb.properties`. Supports dict-like access.

-   Metadata: `title`, `subject`, `creator`, `keywords`, `description`, `last_modified_by`, `category`, `company`.

```python
wb.properties["title"] = "My Report"
print(wb.properties["creator"])
```

---

### `Column` Class
Accessed via `ws.column(col_index)` or `ws.column("A")`.

#### Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `width` | `float` | Get or set the column width. |
| `hidden` | `bool` | Get or set whether the column is hidden. |
| `style_index` | `int` | Get or set the default style index for the column. |

---

### `Formula` Class
Accessed via `cell.formula`.

#### Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `text` | `str` | Get or set the formula string. |

#### Methods

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `clear()` | `None` | Remove the formula from the cell. |

---

### `MergeCells` Class
Accessed via `ws.merges`. Represents the collection of merged ranges in a worksheet.

#### Methods

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `append(reference)` | `None` | Create a merged range (e.g., `"A1:B2"`). |
| `delete(index)` | `None` | Remove a merged range by its index. |
| `find(reference)` | `int` | Find the index of a merged range. Returns -1 if not found. |
| `__len__()` | `int` | Return the number of merged ranges. |
| `__getitem__(index)` | `XLMergeCell` | Get a merged range object by index. |
| `__iter__()` | `Iterator` | Iterate over all merged ranges. |
| `__contains__(ref)` | `bool` | Check if a reference is within any merged range. |

---

### `XLComments` Class
Accessed via `ws._sheet.comments()`.

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `add_author(name)` | `int` | Add a new author to the workbook. |
| `set(ref, text, author_id=0)` | `None` | Set comment for a cell reference. |
| `get(ref_or_idx)` | `str \| XLComment` | Get comment text or object. |
| `shape(cell_ref)` | `XLShape` | Get the VML shape object for the comment box. |
| `count()` | `int` | Number of comments in the sheet. |

---

### `XLShape` Class
Represents the visual properties of a comment box.

#### Methods

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `style()` | `XLShapeStyle` | Access size, position, and visibility properties. |
| `client_data()` | `XLShapeClientData` | Access Excel-specific anchor and auto-fill data. |

---

### `XLShapeStyle` Class

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `set_width(val)` | `None` | Set box width in points. |
| `set_height(val)` | `None` | Set box height in points. |
| `show() / hide()` | `None` | Set comment visibility. |

---

### `XLShapeClientData` Class

| Method | Return Type | Description |
| :--- | :--- | :--- |
| `set_anchor(str)` | `None` | Set box position/size using grid coordinates. |
| `set_auto_fill(bool)`| `None` | Enable/disable automatic box sizing. |

---

### `ImageInfo` Class
Returned by `wb.get_embedded_images()`.

#### Attributes

| Attribute | Type | Description |
| :--- | :--- | :--- |
| `name` | `str` | Filename of the image. |
| `path` | `str` | Internal path in the XLSX archive. |
| `extension` | `str` | File extension (e.g., "png"). |

---

### Helper Functions

#### `load_workbook`
```python
def load_workbook(filename: str) -> Workbook:
    """Open an existing Excel file."""
```

#### `load_workbook_async`
```python
async def load_workbook_async(filename: str) -> Workbook:
    """Asynchronously open an existing Excel file."""
```

#### `is_date_format`
```python
def is_date_format(format_code: int | str) -> bool:
    """
    Check if a number format code (int) or string represents a date/time format.
    Useful for determining if a cell value should be treated as a datetime.
    """
```

---


## Performance

`pyopenxlsx` is built for speed. By leveraging the C++ OpenXLSX engine and providing optimized bulk operations, it significantly outperforms pure-Python alternatives.

### Benchmarks (pyopenxlsx vs openpyxl)

| Scenario | pyopenxlsx | openpyxl | Speedup |
| :--- | :--- | :--- | :--- |
| **Read** (20,000 cells) | **~7.2ms** | ~145ms | **20.2x** |
| **Write** (1,000 cells) | **~4.8ms** | ~8.1ms | **1.7x** |
| **Write** (50,000 cells) | **~169ms** | ~305ms | **1.8x** |
| **Bulk Write** (50,000 cells) | **~74ms** | N/A | **4.1x** |
| **Iteration** (20,000 cells) | **~80ms** | ~150ms | **1.9x** |
| **Bulk Write** (1,000,000 cells) | **~1.5s** | ~6.2s | **4.1x** |

### Resource Usage (1,000,000 cells)

| Library | Execution Time | Memory Delta | CPU Load |
| :--- | :--- | :--- | :--- |
| **pyopenxlsx** | **~1.5s** | ~400 MB | ~99% |
| **openpyxl** | ~6.2s | ~600 MB* | ~99% |

> [!NOTE]
> *Memory delta for `openpyxl` can be misleading due to Python's garbage collection timing during the benchmark. However, `pyopenxlsx` consistently shows lower memory pressure for bulk operations as data is handled primarily in C++.

### Why is it faster?
1. **C++ Foundation**: Core operations happen in highly optimized C++.
2. **Reduced Object Overhead**: `pyopenxlsx` minimizes the creation of many Python `Cell` objects during bulk operations.
3. **Efficient Memory Mapping**: Leverages the memory-efficient design of OpenXLSX.
4. **Asynchronous I/O**: Key operations are available as non-blocking coroutines to maximize throughput in concurrent applications.

---

## Development

### Run Tests

```bash
# Run all tests
uv run pytest

# With coverage
uv run pytest --cov=src/pyopenxlsx --cov-report=term-missing
```

## License

BSD 3-Clause License.
The underlying OpenXLSX library is licensed under the MIT License, and nanobind under a BSD-style license.
