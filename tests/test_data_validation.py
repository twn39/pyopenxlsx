from pyopenxlsx import Workbook, XLDataValidationType, XLDataValidationOperator

def test_list_validation(tmp_path):
    path = tmp_path / "test_dv.xlsx"
    wb = Workbook()
    ws = wb.active
    
    # Method 1: Using high-level add_validation
    dv = ws.data_validations.add_validation(
        "A1", 
        type="list", 
        formula1='"Apple,Banana,Cherry"',
        show_drop_down=True
    )
    
    assert len(ws.data_validations) == 1
    assert dv.type == XLDataValidationType.List
    assert dv.sqref == "A1"
    
    # Method 2: Manual configuration
    dv2 = ws.data_validations.append()
    dv2.sqref = "B1:B10"
    dv2.type = XLDataValidationType.Whole
    dv2.operator = XLDataValidationOperator.Between
    dv2.formula1 = "1"
    dv2.formula2 = "100"
    dv2.set_prompt("Enter Number", "Please enter a number between 1 and 100")
    dv2.set_error("Invalid Number", "The number must be between 1 and 100", style="stop")
    
    assert len(ws.data_validations) == 2
    
    wb.save(path)
    
    # Reload and verify
    wb2 = Workbook(path)
    ws2 = wb2.active
    assert len(ws2.data_validations) == 2
    
    dv_list = ws2.data_validations["A1"]
    assert dv_list.type == XLDataValidationType.List
    
    dv_range = ws2.data_validations["B1:B10"]
    assert dv_range.type == XLDataValidationType.Whole
    assert dv_range.formula1 == "1"
    assert dv_range.formula2 == "100"

def test_validation_iteration():
    wb = Workbook()
    ws = wb.active
    ws.data_validations.add_validation("A1", type="whole", formula1="1", formula2="10")
    ws.data_validations.add_validation("B1", type="decimal", formula1="0.1", formula2="0.9")
    
    sqrefs = [dv.sqref for dv in ws.data_validations]
    assert "A1" in sqrefs
    assert "B1" in sqrefs

def test_validation_removal():
    wb = Workbook()
    ws = wb.active
    ws.data_validations.add_validation("A1", type="list", formula1='"X,Y"')
    ws.data_validations.add_validation("B1", type="list", formula1='"A,B"')
    
    assert len(ws.data_validations) == 2
    ws.data_validations.remove("A1")
    assert len(ws.data_validations) == 1
    assert ws.data_validations[0].sqref == "B1"
    
    ws.data_validations.clear()
    assert len(ws.data_validations) == 0

def test_reference_list(tmp_path):
    path = tmp_path / "test_ref_dv.xlsx"
    wb = Workbook()
    ws = wb.active
    
    # Source data on Sheet2
    ws2 = wb.create_sheet("Source")
    ws2.cell(1, 1, "Red")
    ws2.cell(2, 1, "Green")
    ws2.cell(3, 1, "Blue")
    
    dv = ws.data_validations.append()
    dv.sqref = "A1"
    dv.type = XLDataValidationType.List
    dv.set_reference_drop_list("Source", "A1:A3")
    
    wb.save(path)
    
    wb_read = Workbook(path)
    dv_read = wb_read.active.data_validations["A1"]
    assert dv_read.type == XLDataValidationType.List
    # OpenXLSX usually stores reference lists in formula1 with an '=' prefix and quotes if needed
    f1 = dv_read.formula1
    assert "Source" in f1 and "A1:A3" in f1
