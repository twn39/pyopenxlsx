from pyopenxlsx import Workbook, load_workbook


def test_cell_formula():
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = 10
    ws["A2"].value = 20
    ws["A3"].formula = "=A1+A2"
    assert ws["A3"].formula == "=A1+A2"

    ws["B1"].formula = "=SUM(A1:A2)"
    assert ws["B1"].formula == "=SUM(A1:A2)"


def test_worksheet_metadata():
    wb = Workbook()
    ws = wb.active
    assert ws.max_row == 0
    assert ws.max_column == 0

    ws["C5"].value = "Data"
    assert ws.max_row == 5
    assert ws.max_column == 3


def test_worksheet_append():
    wb = Workbook()
    ws = wb.active
    data = [1, "two", 3.0, True]
    ws.append(data)

    assert ws.max_row == 1
    assert ws.max_column == 4
    assert ws.cell(1, 1).value == 1
    assert ws.cell(1, 2).value == "two"
    assert ws.cell(1, 3).value == 3.0
    assert ws.cell(1, 4).value is True

    ws.append(["Next", "Row"])
    assert ws.max_row == 2
    assert ws.cell(2, 1).value == "Next"


def test_worksheet_rows_iterator():
    wb = Workbook()
    ws = wb.active
    ws.append([1, 2])
    ws.append([3, 4])

    rows = list(ws.rows)
    assert len(rows) == 2
    assert len(rows[0]) == 2
    assert rows[0][0].value == 1
    assert rows[1][1].value == 4


def test_save_load_formulas(tmp_path):
    fn = tmp_path / "formula.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = 100
    ws["A2"].formula = "=A1*2"
    wb.save(str(fn))

    wb2 = load_workbook(str(fn))
    ws2 = wb2.active
    assert ws2["A1"].value == 100
    assert ws2["A2"].formula == "=A1*2"
