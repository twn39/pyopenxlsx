"""Tests for FormulaEngine convenience façade."""

from pyopenxlsx import FormulaEngine, Workbook, calculation_options


def test_evaluate_strips_equals_and_helpers():
    engine = FormulaEngine()
    assert engine.evaluate("=1+2*3") == 7
    assert engine.evaluate("1+2*3") == 7
    assert engine.sum("10, 20, 30") == 60
    assert engine.average("10, 20") == 15
    assert engine.count("10, 20, \"x\"") == 2
    assert engine.if_("1>0", "1", "0") == 1
    assert engine.evaluate_many(["SUM(1,2)", "AVERAGE(2,4)"]) == [3, 3]


def test_contextual_sum():
    wb = Workbook()
    ws = wb.active
    ws.write_row(1, [10, 20, 30])
    engine = FormulaEngine()
    assert engine.sum("A1:C1", ws) == 60
    assert engine.evaluate("SUM(A1:C1)", worksheet=ws) == 60
    wb.close()


def test_calculation_options_helper():
    opts = calculation_options(write_back=True, use_defined_names=True)
    assert opts.write_back is True
    assert opts.use_defined_names is True
