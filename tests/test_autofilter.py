import pytest
from pyopenxlsx import Workbook

def test_autofilter():
    wb = Workbook()
    ws = wb.active
    
    # Setup autofilter range
    ws.auto_filter = "A1:D10"
    
    af = ws.auto_filter
    assert af == "A1:D10"
    assert str(af) == "A1:D10"
    assert bool(af) is True
    
    # Configure a filter column
    fc = af[0]  # First column
    assert fc.col_id == 0
    fc.add_filter("Apple")
    fc.add_filter("Banana")
    
    # Configure custom filter
    fc2 = af[1]
    fc2.set_custom_filter("greaterThan", "10", logic="and", op2="lessThan", val2="20")
    fc2.set_custom_filter("equal", "15") # Without logic
    
    # Configure top10
    fc3 = af[2]
    fc3.set_top10(5, percent=False, top=True)
    fc3.set_top10(10, percent=True, top=False)
    
    # Clear filter
    fc.clear()
    
    ws.apply_auto_filter()
    assert True

