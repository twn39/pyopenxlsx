import xml.etree.ElementTree as ET
from pyopenxlsx import (
    Workbook,
    XLDataValidationType,
    XLIMEMode,
)

def test_data_validation_ooxml_full(tmp_path):
    """
    Comprehensive verification of Data Validation XML generation.
    Checks all major attributes and child elements (formula1, formula2).
    """
    wb = Workbook()
    ws = wb.active
    
    # 1. Complex whole number validation
    dv1 = ws.data_validations.add_validation(
        "A1:A10",
        type="whole",
        operator="not_between",
        formula1="10",
        formula2="20",
        allow_blank=False,
        show_input_message=True,
        show_error_message=True,
        prompt_title="Input Rule",
        prompt="Enter value not between 10 and 20",
        error_title="Invalid Value",
        error="Value must be outside 10-20 range",
    )
    # Set error style manually
    dv1.set_error("Invalid Value", "Value must be outside 10-20 range", style="warning")
    
    # 2. List validation with drop down disabled
    ws.data_validations.add_validation(
        "B1",
        type="list",
        formula1='"Yes,No,Maybe"',
        show_drop_down=False # showDropDown="1" means HIDE the arrow in some versions? 
                             # Actually OOXML showDropDown="1" means DISABLE the dropdown.
                             # Wait, OpenXLSX's setShowDropDown(bool) implementation:
                             # if true, it sets showDropDown="1".
    )
    
    # 3. Decimal validation
    ws.data_validations.add_validation(
        "C1",
        type="decimal",
        operator="greater_than",
        formula1="0.5"
    )

    # Save and inspect XML
    path = tmp_path / "ooxml_dv_full.xlsx"
    wb.save(str(path))
    
    # We use the workbook's internal archive access to read the XML
    xml_data = wb.get_archive_entry("xl/worksheets/sheet1.xml")
    wb.close()
    
    root = ET.fromstring(xml_data)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    
    dv_container = root.find('main:dataValidations', ns)
    assert dv_container is not None
    assert dv_container.get('count') == "3"
    
    dvs = dv_container.findall('main:dataValidation', ns)
    assert len(dvs) == 3
    
    # Verify DV1 (Whole Number, Not Between)
    v1 = next(v for v in dvs if v.get('sqref') == "A1:A10")
    assert v1.get('type') == "whole"
    assert v1.get('operator') == "notBetween"
    assert v1.get('allowBlank') == "0" or v1.get('allowBlank') is None # Default is 1/true if absent
    assert v1.get('showInputMessage') == "1"
    assert v1.get('showErrorMessage') == "1"
    assert v1.get('promptTitle') == "Input Rule"
    assert v1.get('prompt') == "Enter value not between 10 and 20"
    assert v1.get('errorTitle') == "Invalid Value"
    assert v1.get('error') == "Value must be outside 10-20 range"
    assert v1.get('errorStyle') == "warning"
    
    f1 = v1.find('main:formula1', ns)
    f2 = v1.find('main:formula2', ns)
    assert f1.text == "10"
    assert f2.text == "20"
    
    # Verify DV2 (List)
    v2 = next(v for v in dvs if v.get('sqref') == "B1")
    assert v2.get('type') == "list"
    # showDropDown="1" in OOXML means suppress the dropdown arrow
    assert v2.get('showDropDown') == "1"
    
    f2_1 = v2.find('main:formula1', ns)
    assert f2_1.text == '"Yes,No,Maybe"'

    # Verify DV3 (Decimal)
    v3 = next(v for v in dvs if v.get('sqref') == "C1")
    assert v3.get('type') == "decimal"
    assert v3.get('operator') == "greaterThan"
    assert v3.find('main:formula1', ns).text == "0.5"

def test_data_validation_ime_mode(tmp_path):
    """Verify IME mode setting in OOXML."""
    wb = Workbook()
    ws = wb.active
    
    dv = ws.data_validations.append()
    dv.sqref = "A1"
    dv.type = XLDataValidationType.Custom
    dv.ime_mode = XLIMEMode.Hiragana
    
    path = tmp_path / "ooxml_ime.xlsx"
    wb.save(str(path))
    
    xml_data = wb.get_archive_entry("xl/worksheets/sheet1.xml")
    wb.close()
    
    root = ET.fromstring(xml_data)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    v = root.find('.//main:dataValidation', ns)
    assert v.get('imeMode') == "hiragana"
