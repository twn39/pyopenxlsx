from pyopenxlsx import Workbook
from pyopenxlsx.formula_engine import FormulaEngine


def test_formula_engine():
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1).value = 10
    ws.cell(row=1, column=2).value = 20
    ws.cell(row=1, column=3).value = 30

    engine = FormulaEngine()

    # Evaluate with context
    result = engine.evaluate("SUM(A1:C1)", ws)
    assert result == 60

    # Evaluate without context
    result_simple = engine.evaluate("SUM(10, 20, 30)")
    assert result_simple == 60
