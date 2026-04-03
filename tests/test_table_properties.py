import pytest
from pyopenxlsx import Workbook

def test_table_properties():
    wb = Workbook()
    ws = wb.active
    ws.write_rows(1, [["A", "B"], [1, 2]])
    table = ws.add_table("TestTable", "A1:B2")
    
    assert table.name == "TestTable"
    table.name = "NewName"
    assert table.name == "NewName"
    
    table.display_name = "NewDisplay"
    assert table.display_name == "NewDisplay"
    
    table.range = "A1:B3"
    assert table.range == "A1:B3"
    
    table.style = "TableStyleMedium5"
    assert table.style == "TableStyleMedium5"
    
    table.show_row_stripes = False
    assert not table.show_row_stripes
    
    table.show_column_stripes = True
    assert table.show_column_stripes
    
    table.show_first_column = True
    assert table.show_first_column
    
    table.show_last_column = True
    assert table.show_last_column
    
    table.show_totals_row = True
    assert table.show_totals_row
    
    table.append_column("Column C")
    assert table._worksheet == ws

