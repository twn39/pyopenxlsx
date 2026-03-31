import os
import zipfile
import xml.etree.ElementTree as ET
from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLTotalsRowFunction


def test_table_totals_row_and_ooxml(tmp_path):
    filename = str(tmp_path / "table_totals_test.xlsx")

    wb = Workbook()
    ws = wb.active

    # 1. Setup data
    headers = ["ID", "Name", "Score"]
    data = [[1, "Alice", 90], [2, "Bob", 80], [3, "Charlie", 95]]

    ws.write_row(1, headers)
    ws.write_rows(2, data)

    # 2. Add table covering A1:C4
    table = ws.add_table("ScoreTable", "A1:C4")

    # Enable totals row
    table.show_totals_row = True

    # Set total functions AND write cell contents
    id_col = table._table.column("ID")
    id_col.set_totals_row_label("Total:")
    ws.set_cell_value(5, 1, "Total:")  # The totals row is at row 5

    score_col = table._table.column("Score")
    score_col.set_totals_row_function(XLTotalsRowFunction.Average)
    ws.cell(5, 3).formula = "=SUBTOTAL(101,ScoreTable[Score])"

    wb.save(filename)
    wb.close()

    # 3. OOXML Verification
    assert os.path.exists(filename)

    with zipfile.ZipFile(filename, "r") as z:
        # Find the table XML
        table_xml_content = z.read("xl/tables/table1.xml")

        # Parse XML
        root = ET.fromstring(table_xml_content)
        # Namespace for OpenXML spreadsheetml
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        # Check TotalsRow properties on the root <table ...> tag
        assert root.attrib.get("totalsRowCount") == "1", "totalsRowCount should be 1"
        assert root.attrib.get("totalsRowShown") == "1", "totalsRowShown should be 1"

        # *** Critical Bug Fix Verification ***
        # The data range is A1:C4. With totals row, the table ref MUST be A1:C5
        assert root.attrib.get("ref") == "A1:C5", (
            "Table 'ref' should automatically expand to include totals row"
        )

        # The autoFilter ref MUST remain A1:C4
        auto_filter = root.find("main:autoFilter", ns)
        assert auto_filter is not None
        assert auto_filter.attrib.get("ref") == "A1:C4", (
            "autoFilter 'ref' must NOT include totals row"
        )

        # Check Columns
        columns_node = root.find("main:tableColumns", ns)
        assert columns_node is not None
        assert columns_node.attrib.get("count") == "3", (
            "Should automatically have 3 columns"
        )

        columns = columns_node.findall("main:tableColumn", ns)

        # Check ID column (first)
        assert columns[0].attrib.get("name") == "ID"
        assert columns[0].attrib.get("totalsRowLabel") == "Total:"

        # Check Score column (third)
        assert columns[2].attrib.get("name") == "Score"
        assert columns[2].attrib.get("totalsRowFunction") == "average"
