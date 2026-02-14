import pytest
import os
from pyopenxlsx import Workbook

def test_workbook_create_overwrite(tmp_path):
    fn = tmp_path / "existing.xlsx"
    # Create an initial file
    wb = Workbook()
    wb.save(str(fn))
    wb.close()
    
    assert os.path.exists(fn)
    
    # Try to create a new workbook at the same path with force_overwrite=False
    # Note: OpenXLSX should throw an exception if forceOverwrite is false and file exists.
    with pytest.raises(Exception):
        wb2 = Workbook(filename=None, force_overwrite=False)
        # We need to manually call create with the existing path because 
        # Workbook() without filename uses a temp file.
        wb2._doc.create(str(fn), False)

def test_workbook_save_async_overwrite(tmp_path):
    import asyncio
    fn = tmp_path / "save_async_overwrite.xlsx"
    # Create an initial file
    wb = Workbook()
    wb.save(str(fn))
    wb.close()
    
    async def run_test():
        wb2 = Workbook()
        with pytest.raises(Exception):
            await wb2.save_async(str(fn), force_overwrite=False)
        wb2.close()
        
    asyncio.run(run_test())

def test_workbook_save_overwrite(tmp_path):
    fn = tmp_path / "save_overwrite.xlsx"
    # Create an initial file
    wb = Workbook()
    wb["Sheet1"]["A1"].value = "Original"
    wb.save(str(fn))
    wb.close()
    
    assert os.path.exists(fn)
    
    # Try to save another workbook to the same path with force_overwrite=False
    wb2 = Workbook()
    wb2["Sheet1"]["A1"].value = "New"
    with pytest.raises(Exception):
        wb2.save(str(fn), force_overwrite=False)
    wb2.close()

def test_workbook_save_overwrite_true(tmp_path):
    fn = tmp_path / "save_overwrite_true.xlsx"
    # Create an initial file
    wb = Workbook()
    wb["Sheet1"]["A1"].value = "Original"
    wb.save(str(fn))
    wb.close()
    
    assert os.path.exists(fn)
    
    # Should work with force_overwrite=True (default)
    wb2 = Workbook()
    wb2["Sheet1"]["A1"].value = "New"
    wb2.save(str(fn), force_overwrite=True)
    wb2.close()
    
    from pyopenxlsx import load_workbook
    wb3 = load_workbook(str(fn))
    assert wb3["Sheet1"]["A1"].value == "New"
    wb3.close()
