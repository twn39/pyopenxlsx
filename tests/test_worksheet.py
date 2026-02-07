from pyopenxlsx import Workbook


def test_sheet_title():
    wb = Workbook()
    ws = wb.active
    assert ws.title == "Sheet1"
    ws.title = "Renamed"
    assert ws.title == "Renamed"
    assert "Renamed" in wb.sheetnames
    assert "Sheet1" not in wb.sheetnames


def test_create_sheet():
    wb = Workbook()
    ws2 = wb.create_sheet("Sheet2")
    assert ws2.title == "Sheet2"
    assert wb.sheetnames == ["Sheet1", "Sheet2"]

    ws0 = wb.create_sheet("Sheet0", 0)
    assert ws0.title == "Sheet0"
    assert wb.sheetnames == ["Sheet0", "Sheet1", "Sheet2"]


def test_remove_sheet():
    wb = Workbook()
    wb.create_sheet("Sheet2")
    wb.remove(wb["Sheet1"])

    assert wb.sheetnames == ["Sheet2"]
    assert wb.active.title == "Sheet2"


def test_copy_worksheet():
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Data"

    ws_copy = wb.copy_worksheet(ws)
    assert ws_copy.title == "Sheet1 Copy"
    assert ws_copy["A1"].value == "Data"
    assert len(wb) == 2


def test_active_sheet():
    wb = Workbook()
    ws2 = wb.create_sheet("Sheet2")
    wb.active = ws2
    assert wb.active.title == "Sheet2"

    wb.active = wb["Sheet1"]
    assert wb.active.title == "Sheet1"


def test_sheet_visibility():
    wb = Workbook()
    ws = wb.active
    wb.create_sheet("VisibleSheet")  # Ensure there's always one visible sheet

    assert ws.sheet_state == "visible"

    ws.sheet_state = "hidden"
    assert ws.sheet_state == "hidden"

    ws.sheet_state = "very_hidden"
    assert ws.sheet_state == "very_hidden"

    ws.sheet_state = "visible"
    assert ws.sheet_state == "visible"


def test_sheet_index():
    wb = Workbook()
    ws1 = wb.active
    ws2 = wb.create_sheet("Sheet2")

    assert ws1.index == 0
    assert ws2.index == 1

    ws1.index = 1
    assert ws1.index == 1
    assert ws2.index == 0
    assert wb.sheetnames == ["Sheet2", "Sheet1"]
