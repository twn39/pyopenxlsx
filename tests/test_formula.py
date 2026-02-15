import pytest
from pyopenxlsx import Workbook, Formula


@pytest.fixture
def wb(tmp_path):
    wb = Workbook()
    yield wb
    wb.close()  # Ensure handles are closed


def test_formula_api(wb):
    ws = wb.active
    ws["A1"].value = 10
    ws["A2"].value = 20

    # 1. Assign string formula
    cell = ws["A3"]
    cell.formula = "=SUM(A1:A2)"

    # 2. Check return type
    f = cell.formula
    assert isinstance(f, Formula)
    assert str(f) == "=SUM(A1:A2)"
    assert f == "=SUM(A1:A2)"

    # 3. Modify via wrapper
    f.text = "=A1+A2"
    assert cell.formula == "=A1+A2"

    # 4. Assign Formula object to another cell
    cell_b3 = ws["B3"]
    cell_b3.formula = f
    assert str(cell_b3.formula) == "=A1+A2"

    # 5. Clear formula
    cell.formula.clear()
    assert str(cell.formula) == ""
    assert cell.formula.text == ""


def test_formula_repr(wb):
    ws = wb.active
    ws["A1"].formula = "=1+1"
    assert repr(ws["A1"].formula) == "Formula('=1+1')"
