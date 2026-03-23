from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLColorScaleRule, XLColor, XLDataBarRule

def test_conditional_formatting(tmp_path):
    file_path = tmp_path / "test_cf.xlsx"
    
    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, [10, 20, 30])
        
        # Test color scale
        c1 = XLColor(255, 0, 0)
        c2 = XLColor(0, 255, 0)
        rule = XLColorScaleRule(c1, c2)
        ws.add_conditional_formatting("A1:C1", rule)
        
        # Test data bar
        rule2 = XLDataBarRule(XLColor(0, 0, 255), True)
        ws.add_conditional_formatting("A2:C2", rule2)
        
        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        # Read back logic (just checking it survives save/load)
        assert ws.cell(1, 1).value == 10
        assert ws.cell(1, 2).value == 20
        assert ws.cell(1, 3).value == 30
