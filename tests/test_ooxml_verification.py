import xml.etree.ElementTree as ET
from pyopenxlsx import (
    Workbook,
    XLDataValidationType,
    XLDataValidationOperator,
    XLDataValidationErrorStyle,
)

def test_ooxml_elements_verification(tmp_path):
    """
    Directly inspect the generated XML in the XLSX archive to ensure 
    hyperlinks and data validation tags are correctly formed according to OOXML spec.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "ValidationTest"
    
    # 1. Add an external hyperlink
    ws.add_hyperlink("A1", "https://github.com/twn39/OpenXLSX", "OpenXLSX Project")
    
    # 2. Add an internal hyperlink
    ws.add_internal_hyperlink("B1", "ValidationTest!C10", "Jump to C10")
    
    # 3. Add a Data Validation list
    validations = ws.data_validations
    v1 = validations.append()
    v1.sqref = "D1:D10"
    v1.type = XLDataValidationType.List
    v1.set_list(["One", "Two", "Three"])
    v1.set_prompt("Select", "Choose a number")
    
    # 4. Add a Numeric Validation
    v2 = validations.append()
    v2.sqref = "E1:E10"
    v2.type = XLDataValidationType.Whole
    v2.operator = XLDataValidationOperator.Between
    v2.formula1 = "1"
    v2.formula2 = "100"
    
    # Save the workbook to disk
    file_path = tmp_path / "ooxml_verify.xlsx"
    wb.save(str(file_path))
    
    # Extract the worksheet XML directly from the archive using our new API
    # The first sheet is sheet1.xml by default in a new workbook
    xml_data = wb.get_archive_entry("xl/worksheets/sheet1.xml")
    wb.close()
    
    # Parse XML
    root = ET.fromstring(xml_data)
    
    # Namespaces
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    
    # --- Verify Hyperlinks ---
    # OOXML: <hyperlinks><hyperlink ref="A1" r:id="rId1" display="OpenXLSX Project"/></hyperlinks>
    hyperlinks_tag = root.find('main:hyperlinks', ns)
    assert hyperlinks_tag is not None, "Missing <hyperlinks> tag"
    
    all_hyperlinks = hyperlinks_tag.findall('main:hyperlink', ns)
    assert len(all_hyperlinks) == 2
    
    # Check A1 (External)
    h_a1 = next(h for h in all_hyperlinks if h.get('ref') == "A1")
    assert h_a1.get('tooltip') == "OpenXLSX Project"
    # External links use r:id which points to xl/worksheets/_rels/sheet1.xml.rels
    assert '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id' in h_a1.attrib
    
    # Check B1 (Internal)
    h_b1 = next(h for h in all_hyperlinks if h.get('ref') == "B1")
    assert h_b1.get('location') == "ValidationTest!C10"
    assert h_b1.get('display') == "ValidationTest!C10" # OpenXLSX sets display to location for internal
    assert h_b1.get('tooltip') == "Jump to C10"
    
    # --- Verify Data Validations ---
    # OOXML: <dataValidations count="2">...</dataValidations>
    validations_tag = root.find('main:dataValidations', ns)
    assert validations_tag is not None, "Missing <dataValidations> tag"
    assert validations_tag.get('count') == "2"
    
    all_v = validations_tag.findall('main:dataValidation', ns)
    assert len(all_v) == 2
    
    # Check List Validation (D1:D10)
    v_list = next(v for v in all_v if v.get('sqref') == "D1:D10")
    assert v_list.get('type') == "list"
    assert v_list.get('promptTitle') == "Select"
    assert v_list.get('prompt') == "Choose a number"
    
    f1 = v_list.find('main:formula1', ns)
    assert f1 is not None
    assert f1.text == '"One,Two,Three"'  # Literal lists are quoted in OOXML
    
    # Check Whole Number Validation (E1:E10)
    v_whole = next(v for v in all_v if v.get('sqref') == "E1:E10")
    assert v_whole.get('type') == "whole"
    assert v_whole.get('operator') == "between"
    
    wf1 = v_whole.find('main:formula1', ns)
    wf2 = v_whole.find('main:formula2', ns)
    assert wf1.text == "1"
    assert wf2.text == "100"
    
    print("OOXML Element Verification Passed!")
