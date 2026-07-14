# Streams API (High Performance I/O)

For very large datasets, `pyopenxlsx` provides streaming I/O that bypasses dense Cell object graphs. `Worksheet.stream_writer()` / `stream_reader()` return high-level `StreamWriter` / `StreamReader` façades over the native stream types.

**Important:** While a stream writer is active on a worksheet, avoid mixing standard DOM cell writes until the writer is closed.

## Stream Writer

`StreamWriter` appends rows sequentially. Date/datetime values use the same coercion and `Workbook.auto_date_formats` behaviour as bulk worksheet APIs (values are written as Excel serials, with optional style indices).

```python
from datetime import date
from pyopenxlsx import Workbook, Font

with Workbook() as wb:
    ws = wb.active
    bold = wb.add_style(font=Font(bold=True))

    with ws.stream_writer() as writer:
        writer.append_row(["ID", "When", "Name", "Score"])
        writer.append_row([(1, bold), date(2023, 1, 1), ("Alice", bold), 99.9])
        for i in range(1_000_000):
            writer.append_row([i, date(2023, 1, 1), f"User_{i}", 99.9])

    wb.save("large_output.xlsx")
```

Explicit `(value, style_index)` tuples are preserved. Plain values inherit default formatting (or auto date styles when enabled).

### Methods

- `append_row(values, row_opts=None)`
- `set_row(row, start_col, values, row_opts=None)` / `set_row_ref(ref, values, ...)`
- Context manager (`with writer:`) and `close()`
- Properties: `is_active`, `last_row`, `max_column`

## Stream Reader

Iterate rows without loading the full worksheet DOM into Python Cell wrappers.

```python
from pyopenxlsx import Workbook

with Workbook("large_input.xlsx") as wb:
    ws = wb.active
    reader = ws.stream_reader()
    for row_data in reader:
        idx = reader.current_row_index
        # process row_data (list of values)
```

Optional `XLStreamReadOptions` (or kwargs `empty_rows` / `apply_number_formats`) control empty-row policy and number-format application.

## Use cases

- Exporting database query results directly to Excel
- Parsing multi-gigabyte workbooks where a full DOM would OOM
