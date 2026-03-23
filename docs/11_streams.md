# Streams API (High Performance I/O)

For extremely large datasets where allocating millions of `Cell` objects (or even Python lists) simultaneously would consume too much memory, `pyopenxlsx` provides `XLStreamWriter` and `XLStreamReader`.

These classes bypass the standard document model and write/read directly to/from the underlying XML streams on disk, offering the lowest possible memory footprint.

## Stream Writer

`stream_writer` allows you to append rows sequentially. 
**Important:** Once a stream writer is active, you should not use standard cell assignment methods on that worksheet until `writer.close()` is called.

```python
from pyopenxlsx import Workbook

with Workbook("large_output.xlsx") as wb:
    ws = wb.active
    
    # Open a stream writer for this worksheet
    writer = ws.stream_writer()
    
    # Append rows one by one
    writer.append_row(["ID", "Name", "Score"])
    for i in range(1000000):
        # Appends a row immediately to the XML stream
        writer.append_row([i, f"User_{i}", 99.9])
        
    # Close the stream to finalize the XML structure
    writer.close()
```

## Stream Reader

`stream_reader` allows you to iterate through rows sequentially without loading the entire worksheet into memory.

```python
from pyopenxlsx import Workbook

with Workbook("large_input.xlsx") as wb:
    ws = wb.active
    
    # Open a stream reader for this worksheet
    reader = ws.stream_reader()
    
    # Iterate through rows sequentially
    while reader.has_next():
        current_row_idx = reader.current_row()
        row_data = reader.next_row() # Returns a list of values
        
        # Process row_data...
        # print(f"Row {current_row_idx}: {row_data}")
```

## Use Cases
- Exporting database query results directly to Excel.
- Parsing multi-gigabyte `.xlsx` files where loading the DOM would trigger Out-Of-Memory errors.
