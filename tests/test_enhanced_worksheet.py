import pytest
from pyopenxlsx import Workbook, XLPane, XLPaneState
import os

def test_freeze_panes(tmp_path):
    file_path = tmp_path / "test_freeze.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.title = "FreezeTest"
        
        # Write some data
        for r in range(1, 20):
            for c in range(1, 10):
                ws.cell(r, c).value = f"R{r}C{c}"
        
        # 1. Freeze first row and first column
        ws.freeze_panes("B2")
        assert ws.has_panes == True
        
        wb.save(file_path)
    
    # Reload and verify
    with Workbook(file_path) as wb:
        ws = wb.active
        assert ws.has_panes == True
        
        # 2. Clear panes
        ws.clear_panes()
        assert ws.has_panes == False
        wb.save(file_path)

def test_split_panes(tmp_path):
    file_path = tmp_path / "test_split.xlsx"
    with Workbook() as wb:
        ws = wb.active
        
        # Split panes at some pixel-like coordinates
        ws.split_panes(2000, 2000, "C3", active_pane="bottomRight")
        assert ws.has_panes == True
        
        wb.save(file_path)

def test_auto_filter(tmp_path):
    file_path = tmp_path / "test_filter.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, ["ID", "Name", "Score"])
        ws.write_rows(2, [
            [1, "Alice", 90],
            [2, "Bob", 85],
            [3, "Charlie", 95]
        ])
        
        # Set auto filter
        ws.auto_filter = "A1:C4"
        assert ws.auto_filter == "A1:C4"
        
        wb.save(file_path)
    
    # Reload and verify
    with Workbook(file_path) as wb:
        ws = wb.active
        assert ws.auto_filter == "A1:C4"
        
        # Clear filter
        ws.auto_filter = None
        assert ws.auto_filter is None
        wb.save(file_path)

def test_worksheet_zoom(tmp_path):
    file_path = tmp_path / "test_zoom.xlsx"
    with Workbook() as wb:
        ws = wb.active
        
        # Default zoom is 100
        assert ws.zoom == 100
        
        # Set zoom to 150%
        ws.zoom = 150
        assert ws.zoom == 150
        
        wb.save(file_path)
    
    # Reload and verify
    with Workbook(file_path) as wb:
        ws = wb.active
        assert ws.zoom == 150

def test_protection_options(tmp_path):
    file_path = tmp_path / "test_protection.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.protect(
            password="lock",
            insert_columns=True,
            insert_rows=True,
            format_cells=True
        )
        
        p = ws.protection
        assert p["protected"] == True
        assert p["password_set"] == True
        assert p["insert_columns"] == True
        assert p["insert_rows"] == True
        assert p["format_cells"] == True
        assert p["sort"] == False
        
        wb.save(file_path)
    
    # Reload and verify
    with Workbook(file_path) as wb:
        ws = wb.active
        p = ws.protection
        assert p["protected"] == True
        assert p["insert_columns"] == True
        assert p["format_cells"] == True
        
        ws.unprotect()
        assert ws.protection["protected"] == False
