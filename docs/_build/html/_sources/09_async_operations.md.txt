# Async Operations API

To maximize throughput in web servers (like FastAPI or Sanic) or concurrent environments, `pyopenxlsx` exposes `async` versions of all I/O-intensive and computationally heavy methods. They run in a threadpool executor under the hood, ensuring the main asyncio event loop is not blocked.

## Async Context Managers

```python
import asyncio
from pyopenxlsx import Workbook, load_workbook_async

async def process_file():
    # Loading asynchronously
    async with await load_workbook_async("data.xlsx") as wb:
        ws = wb.active
        
        # Read async
        val = await ws.get_cell_value_async(1, 1)
        
        # Write async
        await ws.set_cell_value_async(2, 1, "Updated")
        
        # Bulk write async
        await ws.write_rows_async(3, [[1, 2], [3, 4]])
        
        # Save async
        await wb.save_async("data_updated.xlsx")

asyncio.run(process_file())
```

## Available Async Methods

**Workbook:**
- `await load_workbook_async(filename, password=None)`
- `await wb.save_async(filename, password=None)`
- `await wb.close_async()`
- `await wb.create_sheet_async(title)`
- `await wb.copy_worksheet_async(ws)`
- `await wb.remove_async(ws)`
- `await wb.add_style_async(...)`
- `await wb.extract_images_async(...)`

**Worksheet:**
- `await ws.append_async(data)`
- `await ws.write_row_async(row, data)`
- `await ws.write_rows_async(start_row, data)`
- `await ws.write_range_async(start_row, start_col, data)`
- `await ws.set_cells_async(cells_batch)`
- `await ws.get_cell_value_async(row, col)`
- `await ws.get_row_values_async(row)`
- `await ws.get_range_data_async(r1, c1, r2, c2)`
- `await ws.get_range_values_async(r1, c1, r2, c2)`
- `await ws.get_rows_data_async()`
- `await ws.merge_cells_async(ref)`
- `await ws.unmerge_cells_async(ref)`
- `await ws.protect_async(password, **granular_options)`
- `await ws.unprotect_async()`
- `await ws.add_image_async(path, anchor)`
