# Encryption and Protection API

`pyopenxlsx` provides robust security features, including full support for ECMA-376 Standard and Agile Encryption for entire workbooks, as well as granular protection options for individual worksheets.

## Workbook Encryption (Password Protection)

You can read and write encrypted `.xlsx` files using a password. Under the hood, this uses OpenXLSX's Agile Encryption capabilities.

### Reading Encrypted Workbooks

To open a password-protected file, simply provide the `password` argument to `Workbook` or `load_workbook()`.

```python
from pyopenxlsx import Workbook, load_workbook

# Using the Workbook constructor
wb = Workbook("secure_financials.xlsx", password="my_secret_password")

# Or using the helper function
wb2 = load_workbook("secure_financials.xlsx", password="my_secret_password")

# Async version
# await load_workbook_async("secure_financials.xlsx", password="my_secret_password")
```

### Writing Encrypted Workbooks

To save a workbook with encryption, provide the `password` argument to the `save()` or `save_async()` methods.

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    ws.cell("A1").value = "Confidential Data"
    
    # Save with Agile Encryption
    wb.save("encrypted_output.xlsx", password="super_secure_password")
    
    # Async version
    # await wb.save_async("encrypted_output.xlsx", password="super_secure_password")
```

---

## Worksheet Protection

In addition to encrypting the entire file, you can lock specific worksheets to prevent users from modifying data, while optionally allowing them to perform certain actions (like sorting, filtering, or formatting).

### Protecting a Worksheet

Use the `protect()` method. You can optionally provide a password that the user must enter in Excel to unprotect the sheet.

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    ws.cell("A1").value = "Do not edit this!"
    
    # Protect the sheet with a password, but allow the user to sort and auto-filter
    ws.protect(
        password="sheet_password", 
        sort=True, 
        auto_filter=True,
        format_cells=False # Explicitly disable formatting (default is False anyway)
    )
    
    wb.save("protected_sheet.xlsx")
```

### Granular Protection Options

The `protect()` method accepts many boolean keyword arguments to fine-tune what a user is allowed to do on a protected sheet:

- `sheet` (default `True`): Enable sheet protection.
- `objects`: Protect drawing objects.
- `scenarios`: Protect scenarios.
- `format_cells`: Allow formatting cells.
- `format_columns`: Allow formatting columns.
- `format_rows`: Allow formatting rows.
- `insert_columns`: Allow inserting columns.
- `insert_rows`: Allow inserting rows.
- `insert_hyperlinks`: Allow inserting hyperlinks.
- `delete_columns`: Allow deleting columns.
- `delete_rows`: Allow deleting rows.
- `sort`: Allow sorting.
- `auto_filter`: Allow using AutoFilters.
- `pivot_tables`: Allow using PivotTables.
- `select_locked_cells` (default `True`): Allow selecting locked cells.
- `select_unlocked_cells` (default `True`): Allow selecting unlocked cells.

### Unprotecting a Worksheet

To remove protection from a worksheet via Python:

```python
ws.unprotect()
```