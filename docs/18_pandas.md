# Pandas Integration API

`pyopenxlsx` offers deep and native support for Python's most popular data analysis library: `pandas`.

By combining the C++ memory-mapped engine of `pyopenxlsx` with the vectorized operations of `pandas`, you can achieve unprecedented speeds when importing and exporting huge datasets (e.g., writing 1 million styled rows in under 8 seconds).

## Exporting a DataFrame (Writing to Excel)

Use `Worksheet.write_dataframe()` to instantly dump a `pandas.DataFrame` into an Excel sheet. 

```python
import pandas as pd
import datetime
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

df = pd.DataFrame({
    "ID": [101, 102, 103],
    "Department": ["Sales", "Engineering", "HR"],
    "Salary": [85000.50, 120000.00, 75000.75],
    "Hire Date": [datetime.date(2021, 5, 1), datetime.date(2019, 8, 15), datetime.date(2022, 1, 10)]
})

# Basic export (writes headers automatically)
ws.write_dataframe(df, start_row=1, start_col=1)

wb.save("pandas_export.xlsx")
```

### High-Performance Column Styling

A common pain point when exporting DataFrames is applying Excel formatting (like Currency `$` or Date `yyyy-mm-dd`) to specific columns without looping over millions of cells in Python (which is extremely slow).

`pyopenxlsx` solves this with the `column_styles` parameter. When provided, the engine automatically switches to the C++ `XLStreamWriter`, injecting your styles natively during the streaming process with **Zero Overhead**.

```python
from pyopenxlsx import Workbook

wb = Workbook()
ws = wb.active

# 1. Create the styles you need
currency_style = wb.add_style(number_format="$#,##0.00")
date_style = wb.add_style(number_format="yyyy-mm-dd")

# 2. Export and style in ONE operation!
# You can map by column name or 0-based column index
ws.write_dataframe(df, column_styles={
    "Salary": currency_style,
    "Hire Date": date_style
})

wb.save("styled_export.xlsx")
```
*Note: Using `column_styles` is capable of styling 1,000,000 rows in just ~3 seconds!*

## Importing to a DataFrame (Reading from Excel)

To extract data back into a `pandas.DataFrame` with maximum performance, use `Worksheet.read_dataframe()`.

Instead of allocating Python `Cell` objects for the entire document, this method uses the underlying C++ XML `XLStreamReader` to extract raw values instantly.

```python
from pyopenxlsx import load_workbook
import pandas as pd

wb = load_workbook("styled_export.xlsx")
ws = wb.active

# 1. Read the data directly into a DataFrame
# You can specify the exact bounding box, or let it read the entire used range
df_read = ws.read_dataframe(header=True)

# 2. Restore Date formats
# Since stream reading extracts raw Excel serial numbers (e.g., 44317.0) for maximum speed,
# you should use pandas' native vectorized functions to restore datetime objects:
if "Hire Date" in df_read.columns:
    df_read["Hire Date"] = pd.to_datetime(
        df_read["Hire Date"], 
        unit='D', 
        origin='1899-12-30'
    ).dt.date

print(df_read)
```

## Async Pandas Operations

If you are building a web backend (like FastAPI) that exports or imports reports, you can use the `async` variants to completely offload the CPU-bound conversions and disk I/O to a threadpool, keeping your event loop responsive.

```python
import asyncio
from pyopenxlsx import Workbook

async def export_report(df):
    wb = Workbook()
    ws = wb.active
    
    # Non-blocking DataFrame write
    await ws.write_dataframe_async(df)
    
    # Non-blocking zip compression and save
    await wb.save_async("async_report.xlsx")
```