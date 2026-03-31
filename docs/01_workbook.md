# Workbook API

The `Workbook` class is the main entry point for creating, modifying, and saving Excel files in `pyopenxlsx`.

## Creating and Loading

### `Workbook(filename=None, force_overwrite=True)`
Creates a new workbook or opens an existing one.
- **Parameters:**
  - `filename` (`str`, optional): Path to an existing `.xlsx` file. If `None`, creates a blank workbook.
  - `force_overwrite` (`bool`): If `True`, allows overwriting existing files when saving.
- **Example:**
  ```python
  from pyopenxlsx import Workbook
  wb = Workbook() # New
  wb_existing = Workbook("data.xlsx") # Load existing
  ```

### `load_workbook(filename)`
Alternative function to load a workbook.
- **Parameters:** `filename` (`str`)
- **Returns:** `Workbook`

### `load_workbook_async(filename)`
Asynchronous version of `load_workbook`.

---

## Properties

### `active`
- **Type:** `Worksheet`
- **Description:** Get or set the currently active worksheet.

### `sheetnames`
- **Type:** `list[str]`
- **Description:** Returns a list of all worksheet names in the workbook.

### `has_macro`
- **Type:** `bool`
- **Description:** Returns `True` if the loaded document contains a VBA macro project (e.g., `vbaProject.bin`). Note: Saving changes to a `.xlsm` file preserves macros losslessly.

### `properties`
- **Type:** `DocumentProperties`
- **Description:** Access standard document properties like `title`, `creator`, `subject`, etc.
  ```python
  wb.properties.title = "My Report"
  ```

### `custom_properties`
- **Type:** `CustomProperties`
- **Description:** Dictionary-like access to custom document properties.
  ```python
  wb.custom_properties["Version"] = "1.0"
  ```

### `defined_names`
- **Type:** `XLDefinedNames`
- **Description:** Manage named ranges (Defined Names) across the workbook.
  ```python
  wb.defined_names.append("GlobalTotal", "Sheet1!$A$1")
  ```

---

## Methods

### `save(filename=None, force_overwrite=True)`
Saves the workbook to disk.
- **Parameters:**
  - `filename` (`str`, optional): The path to save to. If `None`, saves over the original file.

### `save_async(filename=None, force_overwrite=True)`
Asynchronously saves the workbook.

### `close()` / `close_async()`
Releases the underlying C++ resources. Recommended to use the workbook as a context manager (`with Workbook() as wb:`) to handle this automatically.

### `create_sheet(title=None, index=None) -> Worksheet`
Creates a new worksheet.
- **Parameters:**
  - `title` (`str`, optional): The name of the new sheet.
  - `index` (`int`, optional): The 0-based position to insert the sheet.

### `remove(worksheet)`
Removes a worksheet from the workbook.

### `copy_worksheet(from_worksheet) -> Worksheet`
Creates a duplicate of an existing worksheet.

### `add_style(...) -> int`
Registers a new cell style in the workbook and returns its integer index.
- **Parameters:** `font`, `fill`, `border`, `alignment`, `number_format`, `protection`
- **Returns:** `int` (Style ID)

### Advanced/Internal Methods

- **`get_embedded_images() -> list[ImageInfo]`**: Gets a list of all images embedded in the workbook archive.
- **`get_image_data(name_or_path: str) -> bytes`**: Gets the binary data for an embedded image.
- **`extract_images(out_dir: str) -> list[str]`**: Extracts all embedded images to the given directory.
- **`get_archive_entries() -> list[str]`**: Lists all files within the underlying `.xlsx` zip archive.
- **`has_archive_entry(path: str) -> bool`**: Checks if a specific file exists within the archive.
- **`get_archive_entry(path: str) -> bytes`**: Reads the raw binary content of a file within the archive.

### Advanced Properties
- **`styles`**: Access the underlying `XLStyles` object.
- **`workbook`**: Access the underlying C++ `XLWorkbook` object.

## Advanced Example: Modifying Document Metadata
```python
from pyopenxlsx import load_workbook

# Use context manager to ensure proper cleanup of C++ bindings
with load_workbook("existing.xlsx") as wb:
    # Update standard properties
    wb.properties.title = "Q4 Financial Report"
    wb.properties.creator = "Finance Bot v2"
    
    # Iterate and print existing custom properties
    print("Previous Custom Properties:")
    for key, value in wb.custom_properties.items():
        print(f"  {key}: {value}")
        
    # Set a new custom property
    wb.custom_properties["Approval_Status"] = "Pending"
    
    # Extract embedded images and zip contents
    images = wb.get_embedded_images()
    if images:
        print(f"Found {len(images)} images in this workbook.")
        
    wb.save("existing_updated.xlsx")
```
