<div align="center">

# PyOpenXLSX

[![PyPI version](https://img.shields.io/pypi/v/pyopenxlsx.svg)](https://pypi.org/project/pyopenxlsx/)
[![Python versions](https://img.shields.io/pypi/pyversions/pyopenxlsx.svg)](https://pypi.org/project/pyopenxlsx/)
[![Downloads](https://img.shields.io/pypi/dm/pyopenxlsx)](https://pypi.org/project/pyopenxlsx/)
[![Build Status](https://github.com/twn39/pyopenxlsx/actions/workflows/build.yml/badge.svg)](https://github.com/twn39/pyopenxlsx/actions/workflows/build.yml)
[![Docs Status](https://github.com/twn39/pyopenxlsx/actions/workflows/docs.yml/badge.svg)](https://twn39.github.io/pyopenxlsx/)
[![Codecov](https://img.shields.io/codecov/c/github/twn39/pyopenxlsx)](https://codecov.io/gh/twn39/pyopenxlsx)
[![License](https://img.shields.io/pypi/l/pyopenxlsx.svg)](https://opensource.org/licenses/BSD-3-Clause)

</div>

> [!IMPORTANT]
> `pyopenxlsx` uses [OpenXLSX-NX](https://github.com/twn39/OpenXLSX-NX) (v1.0.0+), a specialized C++ fork that includes critical performance optimizations and functional enhancements (such as agile encryption, streaming I/O, vector shapes, threaded comments, and custom properties) not currently available in the upstream repository.

`pyopenxlsx` is a high-performance Python binding for the [OpenXLSX-NX](https://github.com/twn39/OpenXLSX-NX) C++ library. It aims to provide significantly faster read/write speeds compared to pure Python libraries like `openpyxl`, while maintaining a Pythonic API design.

## Core Features

-   **High Performance**: Powered by the modern C++17 OpenXLSX-NX library.
-   **Pythonic API**: Intuitive interface with properties, iterators, and context managers.
-   **Streaming I/O**: Bypass the DOM entirely with `XLStreamWriter` and `XLStreamReader` for memory-efficient bulk data processing.
-   **Security**: Full support for ECMA-376 Standard and Agile Encryption (read/write password-protected files) and granular worksheet protection.
-   **Async Support**: `async/await` support for key I/O operations.
-   **Rich Styling**: Comprehensive support for fonts, fills, borders, alignments, and number formats.
-   **Extended Metadata**: Support for both standard and **custom document properties**.
-   **Advanced Content**: Support for **images**, **vector shapes**, **hyperlinks** (external/internal), and modern **threaded comments**.
-   **Memory Safety**: Combines C++ efficiency with Python's automatic memory management.

## Tech Stack

| Component | Technology |
| :--- | :--- |
| **C++ Core** | [OpenXLSX-NX](https://github.com/twn39/OpenXLSX-NX) |
| **Bindings** | [nanobind](https://github.com/wjakob/nanobind) |
| **Build System** | [scikit-build-core](https://github.com/scikit-build/scikit-build-core) & [CMake](https://cmake.org/) |

## pyopenxlsx vs openpyxl: Feature Comparison

While `openpyxl` is a great pure-Python library, `pyopenxlsx` is designed to solve critical performance bottlenecks and add modern enterprise features by leveraging a C++ engine.

| Feature / Capability | `pyopenxlsx` (OpenXLSX-NX) | `openpyxl` | Notes |
| :--- | :--- | :--- | :--- |
| **Underlying Engine** | **C++17** (`nanobind` wrapped) | Pure Python | `pyopenxlsx` is heavily optimized for low-level memory management. |
| **Execution Speed** | **Extremely Fast** (Up to 160x) | Slower | Pure Python loop overhead makes parsing large files sluggish. |
| **Memory Footprint** | **Minimal** (C++ Memory Mapping) | High | Parsing large files in `openpyxl` often leads to OOM errors. |
| **Asyncio Support** | ✅ **Native** (`await load_workbook_async`) | ❌ No | `pyopenxlsx` offloads heavy I/O to a threadpool, perfect for Web APIs (FastAPI/Django). |
| **Agile Encryption (Passwords)** | ✅ **Native Read & Write** | ❌ No | `openpyxl` cannot read/write password-protected `.xlsx` files without 3rd-party decryption tools. |
| **Threaded Comments** | ✅ **Full Support** (Conversations/Replies) | ❌ No / Can be lost | `pyopenxlsx` supports modern Excel conversational comments and resolution states. |
| **Vector Shapes** | ✅ **Native Support** (20+ Shapes) | ❌ No | Draw complex vector shapes (Arrows, Flowcharts, etc.) directly. |
| **Formula Evaluation** | ✅ **Built-in C++ Engine** | ❌ No | `pyopenxlsx` can statically evaluate simple formulas without Excel installed. |
| **Streaming I/O** | ✅ **Direct to disk with Styles** | ⚠️ Partial (WriteOnly) | `pyopenxlsx` can stream styled data directly to the archive, bypassing the DOM. |
| **Granular Sheet Protection** | ✅ **Deep Control** (20+ specific flags) | ✅ Yes | `pyopenxlsx` exposes extensive ECMA-376 locking options. |
| **Styles Architecture** | ✅ **Declarative (Index-based)** | ⚠️ Object-based | `pyopenxlsx` reuses style indices, saving massive amounts of memory on huge datasets. |
| **Charts** | ⚠️ Basic (Bar, Line, etc.) | ✅ **Highly Advanced** | `openpyxl` currently has more mature support for extremely complex/3D charts. |
| **Environment** | Pre-compiled Wheels required | Any Python env | `pyopenxlsx` provides wheels for major OS/Architectures via CI. |

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

### Pivot Tables

Create dynamic pivot tables based on worksheet data.

```python
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLPivotTableOptions, XLPivotField, XLPivotSubtotal

with Workbook() as wb:
    # 1. Write source data
    ws = wb.active
    ws.title = "SalesData"
    ws.write_row(1, ["Region", "Product", "Sales"])
    ws.write_rows(2, [
        ["North", "Apples", 100],
        ["South", "Bananas", 300],
        ["North", "Oranges", 150]
    ])
    
    # 2. Create a separate sheet for the Pivot Table
    ws_pivot = wb.create_sheet("PivotSheet")
    
    # 3. Configure options
    options = XLPivotTableOptions()
    options.name = "SalesPivot"
    options.source_range = "SalesData!A1:C4"
    options.target_cell = "A3" # Note: Target cell must NOT include sheet name
    
    # 4. Define fields
    r = XLPivotField()
    r.name = "Region"
    r.subtotal = XLPivotSubtotal.Sum
    options.rows = [r]

    c = XLPivotField()
    c.name = "Product"
    c.subtotal = XLPivotSubtotal.Sum
    options.columns = [c]

    d = XLPivotField()
    d.name = "Sales"
    d.subtotal = XLPivotSubtotal.Sum
    d.custom_name = "Total Sales"
    options.data = [d]
    
    # 5. Add to the new sheet
    ws_pivot._sheet.add_pivot_table(options)
    
    wb.save("pivot.xlsx")
```

### Insert Images and Vector Shapes

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# 1. Insert image at A1, automatically maintaining aspect ratio
# Requires Pillow: pip install pillow
ws.add_image("logo.png", anchor="A1", width=200)

# 2. Or specify exact dimensions
ws.add_image("banner.jpg", anchor="B5", width=400, height=100)

# 3. Add Native Vector Shapes
ws.add_shape(
    row=2, col=5, shape_type="Arrow", 
    name="MyArrow", text="Point!", 
    fill_color="FF0000", line_width=2.5,
    rotation=90
)

wb.save("media.xlsx")
```

### Comments & Threaded Replies

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# 1. Simple or multiline legacy comments
ws["A1"].comment = "Short comment"

# 2. Modern Threaded Comments (Conversations)
author_id = wb._doc.persons().add_person("Curry Tang")
threads = ws._sheet.threaded_comments()

root_comment = threads.add_comment("B2", author_id, "Please review this cell.")
threads.add_reply(root_comment.id(), author_id, "Fixed!")

wb.save("comments.xlsx")
```

---


### Conditional Formatting

Highlight specific data using visual rules like color scales and data bars.

```python
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLColorScaleRule, XLDataBarRule, XLColor

wb = Workbook()
ws = wb.active
ws.write_rows(1, [[1, 2, 3], [4, 5, 6], [7, 8, 9]])

# 1. Color Scale Rule (Red to Green)
scale_rule = XLColorScaleRule(XLColor(255, 0, 0), XLColor(0, 255, 0))
ws.add_conditional_formatting("A1:C1", scale_rule)

# 2. Data Bar Rule (Blue bars)
bar_rule = XLDataBarRule(XLColor(0, 0, 255), show_value_text=True)
ws.add_conditional_formatting("A2:C2", bar_rule)

wb.save("conditional_formatting.xlsx")
```

### High Performance Streams (Low Memory I/O)

For writing massive datasets without consuming memory for Python objects, use the direct stream writer.

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    
    # Open a direct XML stream writer
    writer = ws.stream_writer()
    
    writer.append_row(["ID", "Timestamp", "Value"])
    for i in range(1_000_000):
        # Writes directly to disk/archive; highly memory efficient
        writer.append_row([i, "2023-01-01", 99.9])
        
    writer.close()
    wb.save("massive_data.xlsx")
```

## API Documentation

The full API documentation has been split into individual modules for easier reading. Please refer to the `docs/` directory:

- [Workbook API](docs/01_workbook.md)
- [Worksheet API](docs/02_worksheet.md)
- [Cell & Range API](docs/03_cell_range.md)
- [Styles API](docs/04_styles.md)
- [Data Validation API](docs/05_data_validation.md)
- [Tables (ListObjects) API](docs/06_tables.md)
- [Pivot Tables API](docs/07_pivot_tables.md)
- [Rich Text API](docs/08_rich_text.md)
- [Async Operations API](docs/09_async_operations.md)
- [Conditional Formatting API](docs/10_conditional_formatting.md)
- [Streams I/O API](docs/11_streams.md)
- [Charts API](docs/12_charts.md)
- [Page Setup & Printing API](docs/13_page_setup.md)
- [Images & Shapes API](docs/14_images_and_shapes.md)
- [Formula Engine API](docs/15_formula_engine.md)
- [Comments & Threaded Comments API](docs/16_comments.md)
- [Encryption & Protection API](docs/17_encryption.md)

---

## Performance

`pyopenxlsx` is built for speed. By leveraging the C++ OpenXLSX-NX engine and providing optimized bulk operations, it significantly outperforms pure-Python alternatives.

### Benchmarks (pyopenxlsx vs openpyxl)

| Scenario | pyopenxlsx | openpyxl | Speedup |
| :--- | :--- | :--- | :--- |
| **Load File** (20,000 cells) | **~0.95ms** | ~154.3ms | **162x** |
| **Single Read** (1 cell in large doc) | **~2.6ms** | ~139.8ms | **53.7x** |
| **Bulk Read** (20,000 cells via values_only) | **~7.8ms** | ~131.6ms | **16.8x** |
| **Write Small** (1,000 cells) | **~4.7ms** | ~8.3ms | **1.7x** |
| **Write Large** (50,000 cells) | **~124.4ms** | ~338.3ms | **2.7x** |
| **Bulk Write Large** (50,000 cells, numpy/range) | **~39.8ms** | N/A | **8.5x** |
| **Extreme Write** (1,000,000 cells) | **~1,539ms** | ~6,635ms | **4.3x** |
| **Bulk Write Extreme** (1,000,000 cells, numpy) | **~649ms** | N/A | **10.2x** |

### Resource Usage (1,000,000 cells)

| Library | Execution Time | Memory Delta | CPU Load |
| :--- | :--- | :--- | :--- |
| **pyopenxlsx** (bulk write) | **~0.65s** | ~200 MB | ~99% |
| **openpyxl** | ~6.6s | ~600 MB* | ~99% |

> [!NOTE]
> *Memory delta for `openpyxl` can be misleading due to Python's garbage collection timing during the benchmark. However, `pyopenxlsx` consistently shows lower memory pressure for bulk operations as data is handled primarily in C++.

### Why is it faster?
1. **C++ Foundation**: Core operations happen in highly optimized C++.
2. **Reduced Object Overhead**: `pyopenxlsx` minimizes the creation of many Python `Cell` objects during bulk operations.
3. **Efficient Memory Mapping**: Leverages the memory-efficient design of OpenXLSX-NX.
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
The underlying OpenXLSX-NX library is licensed under the MIT License, and nanobind under a BSD-style license.
