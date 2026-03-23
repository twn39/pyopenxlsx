from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLSparklineType

def test_worksheet_extras(tmp_path):
    file_path = tmp_path / "test_extras.xlsx"
    
    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, ["Data", 10, 20, 30, 40])
        
        # Auto fit column
        ws.auto_fit_column(1)
        
        # Auto filter (using the method directly rather than property)
        ws.apply_auto_filter()
        
        # Set print area
        ws.set_print_area("A1:E10")
        
        # Set print titles
        ws.set_print_title_rows(1, 2)
        ws.set_print_title_cols(1, 2)
        
        # Add sparkline
        ws.add_sparkline("F1", "B1:E1", XLSparklineType.Line)
        
        # Add comment
        ws.add_comment("A1", "This is a test comment", "Author")
        
        # Add table
        table = ws.add_table("MyTable", "A1:E2")
        assert table is not None
        
        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        # Basic sanity check
        assert ws.cell(1, 1).value == "Data"
        # Check table
        assert len(ws.tables) >= 1

