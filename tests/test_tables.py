from pyopenxlsx import Workbook


def test_table_creation(tmp_path):
    path = tmp_path / "test_table.xlsx"
    wb = Workbook()
    ws = wb.active

    # Write some data for the table
    headers = ["ID", "Name", "Score"]
    ws.write_row(1, headers)
    ws.write_row(2, [1, "Alice", 95])
    ws.write_row(3, [2, "Bob", 88])

    # Get the table object (it is created automatically if it doesn't exist)
    table = ws.table
    table.name = "ScoreTable"
    table.range = "A1:C3"
    table.style = "TableStyleMedium2"
    table.show_row_stripes = True
    table.show_column_stripes = False

    assert table.name == "ScoreTable"
    assert table.range == "A1:C3"
    assert table.style == "TableStyleMedium2"

    wb.save(path)

    # Reload and verify
    wb2 = Workbook(path)
    ws2 = wb2.active
    table2 = ws2.table
    assert table2.name == "ScoreTable"
    assert table2.range == "A1:C3"
    assert table2.style == "TableStyleMedium2"
    assert table2.show_row_stripes is True


def test_table_append_column():
    wb = Workbook()
    ws = wb.active
    table = ws.table
    table.range = "A1:B2"

    table.append_column("NewCol")
    # OpenXLSX doesn't currently provide a way to list columns,
    # but we can verify it doesn't crash.
    # We could check the XML if we wanted to be sure.
    assert table.name.startswith("Table")
