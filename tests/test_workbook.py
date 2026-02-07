import os
from pyopenxlsx import Workbook, load_workbook


def test_new_workbook():
    wb = Workbook()
    assert len(wb) == 1
    assert "Sheet1" in wb.sheetnames
    assert wb.active.title == "Sheet1"


def test_workbook_save_load(tmp_path):
    fn = tmp_path / "test.xlsx"
    wb = Workbook()
    wb["Sheet1"]["A1"].value = "Hello World"
    wb.save(str(fn))
    wb.close()

    assert os.path.exists(fn)

    wb2 = load_workbook(str(fn))
    assert len(wb2) == 1
    assert wb2["Sheet1"]["A1"].value == "Hello World"
    wb2.close()


def test_workbook_iteration():
    wb = Workbook()
    wb.create_sheet("Sheet2")
    wb.create_sheet("Sheet3")

    names = [ws.title for ws in wb]
    assert names == ["Sheet1", "Sheet2", "Sheet3"]
    assert len(wb) == 3


def test_workbook_contains():
    wb = Workbook()
    assert "Sheet1" in wb
    assert "NonExistent" not in wb
