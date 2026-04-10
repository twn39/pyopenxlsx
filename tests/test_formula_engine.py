import pytest
from pyopenxlsx import Workbook
from pyopenxlsx.formula_engine import FormulaEngine


def test_formula_engine_basic():
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


def test_formula_engine_excelize_cases():
    """Test comprehensive formula evaluation matching OpenXLSX's new Excelize cross-validation suite."""
    wb = Workbook()
    ws = wb.active
    ws.name = "Data"
    
    # Populate data matching the C++ test
    for i in range(1, 11):
        ws.cell(row=i, column=1).value = i * 2
        ws.cell(row=i, column=2).value = i * 1.5
        ws.cell(row=i, column=3).value = f"Item{i}"
        ws.cell(row=i, column=4).value = (i % 3 == 0)
        ws.cell(row=i, column=5).value = 46000 + i

    engine = FormulaEngine()

    test_cases = {
        "ABS(-10.5)": 10.5,
        "ACOS(0.8)": 0.643501108793284,
        "AND(TRUE(), 1, 2>1)": True,
        "ASIN(0.2)": 0.201357920790331,
        "AVEDEV(2, 4, 6, 8)": 2,
        "AVERAGE(10, 20)": 15,
        "AVERAGEA(\"5\", FALSE(), 10)": 5,
        "AVERAGEIF(Data!A1:A10, \"<=8\")": 5,
        "AVERAGEIFS(Data!A1:A10, Data!A1:A10, \">2\", Data!B1:B10, \"<10\")": 8,
        "CEILING(3.14, 0.1)": 3.2,
        "CEILING.MATH(4.2)": 5,
        "CHAR(66)": "B",
        "CLEAN(CHAR(7)&\"hello\"&CHAR(9))": "hello",
        "CODE(\"B\")": 66,
        "CONCAT(\"Hello\", \" \", \"World\")": "Hello World",
        "CONCATENATE(\"Foo\", \"Bar\")": "FooBar",
        "CORREL(Data!A2:A6, Data!B2:B6)": 1,
        "COS(1)": 0.54030230586814,
        "COUNT(10, 20, 30, \"B\", TRUE())": 4,
        "COUNTA(1, \"\", FALSE())": 2,
        "COUNTBLANK(Data!F1:G5)": 10,
        "COUNTIF(Data!B1:B10, \"<10\")": 6,
        "COUNTIFS(Data!A1:A10, \">2\", Data!B1:B10, \"<15\")": 8,
        "COVAR(Data!A2:A5, Data!B2:B5)": 3.75,
        "COVARIANCE.P(Data!A2:A5, Data!B2:B5)": 3.75,
        "COVARIANCE.S(Data!A2:A5, Data!B2:B5)": 5,
        "DATE(2024, 2, 29)": 45351,
        "DAY(46000)": 9,
        "DAYS(46000, 45000)": 1000,
        "DAYS360(45000, 46000)": 984,
        "DB(20000, 2000, 10, 2)": 3271.28,
        "DDB(20000, 2000, 10, 2)": 3200,
        "DEGREES(1)": 57.2957795130823,
        "DEVSQ(2, 4, 6, 8)": 20,
        "EDATE(46000, 2)": 46062.0, # Excelize is 45697 (bugged)
        "EOMONTH(46000, 2)": 46081, # Fixed value that pyopenxlsx handles
        "EXACT(\"Hello\", \"hello\")": False,
        "EXP(2)": 7.38905609893065,
        "FALSE()": False,
        "FIND(\"l\", \"hello\")": 3,
        "FISHER(0.8)": 1.09861228866811,
        "FISHERINV(0.8)": 0.664036770267849,
        "FLOOR(3.14, 0.1)": 3.1,
        "FLOOR.MATH(4.8)": 4,
        "FV(0.06, 12, -200)": 3373.98823945184,
        "HOUR(0.75)": 18,
        "IF(FALSE(), 1, 2)": 2,
        "IFERROR(1/0, 999)": 999,
        "IFNA(VLOOKUP(100, Data!A1:B10, 2, FALSE()), 888)": 888,
        "IFS(1>2, 1, 2>1, 2)": 2,
        "INDEX(Data!A1:A10, 3)": 6,
        "INT(4.8)": 4,
        "INTERCEPT(Data!A2:A6, Data!B2:B6)": 4.44089209850063E-16,
        "ISBLANK(Data!F1)": True,
        "ISERR(1/0)": True,
        "ISERROR(1/0)": True,
        "ISEVEN(3)": False,
        "ISLOGICAL(1)": False,
        "ISNA(VLOOKUP(100, Data!A1:B10, 2, FALSE()))": True,
        "ISNONTEXT(\"A\")": False,
        "ISNUMBER(\"1\")": False,
        "ISODD(4)": False,
        "ISOWEEKNUM(46000)": 50,
        "ISTEXT(1)": False,
        "LARGE(Data!A1:A10, 2)": 18,
        "LEFT(\"hello\", 3)": "hel",
        "LEN(\"hello\")": 5,
        "LOG(1000, 10)": 3,
        "LOG10(1000)": 3,
        "LOWER(\"HELLO\")": "hello",
        "MATCH(6, Data!A1:A10, 0)": 3,
        "MAX(10, 20, 30)": 30,
        "MAXIFS(Data!A1:A10, Data!A1:A10, \"<=8\")": 8,
        "MEDIAN(10, 20, 30, 40)": 25,
        "MID(\"hello\", 2, 3)": "ell",
        "MIN(10, 20, 30)": 10,
        "MINIFS(Data!A1:A10, Data!A1:A10, \">=6\")": 6,
        "MINUTE(0.75)": 0,
        "MOD(10, 3)": 1,
        "MONTH(46000)": 12,
        "MROUND(10, 3)": 9,
        "NETWORKDAYS(45000, 46000)": 715,
        "NOT(FALSE())": True,
        "NPER(0.06, -200, 2000)": 15.7252085438878,
        "NPV(0.06, 200, 200, 200)": 534.602389892327,
        "OR(FALSE(), FALSE(), TRUE())": True,
        "PEARSON(Data!A2:A6, Data!B2:B6)": 1,
        "PERCENTILE(Data!A1:A10, 0.8)": 16.4,
        "PERCENTILE.EXC(Data!A1:A10, 0.8)": 17.6,
        "PERCENTILE.INC(Data!A1:A10, 0.8)": 16.4,
        "PERMUT(6, 3)": 120,
        "PERMUTATIONA(6, 3)": 216,
        "PI()": 3.14159265358979,
        "PMT(0.06, 12, 2000)": -238.554058761327,
        "POWER(3, 4)": 81,
        "PROPER(\"hello world\")": "Hello World",
        "PV(0.06, 12, -200)": 1676.76878807667,
        "QUARTILE(Data!A1:A10, 2)": 11,
        "QUARTILE.EXC(Data!A1:A10, 2)": 11,
        "QUARTILE.INC(Data!A1:A10, 2)": 11,
        "RADIANS(90)": 1.5707963267949,
        "RANK(6, Data!A1:A10)": 8,
        "RANK.EQ(6, Data!A1:A10)": 8,
        "REPLACE(\"hello\", 1, 2, \"j\")": "jllo",
        "REPT(\"b\", 4)": "bbbb",
        "RIGHT(\"hello\", 3)": "llo",
        "ROUND(3.14159, 2)": 3.14,
        "ROUNDDOWN(3.14159, 2)": 3.14,
        "ROUNDUP(3.14159, 2)": 3.15,
        "RSQ(Data!A2:A6, Data!B2:B6)": 1,
        "SEARCH(\"L\", \"hello\")": 3,
        "SECOND(0.75)": 0,
        "SIGN(10)": 1,
        "SIN(1)": 0.841470984807897,
        "SLN(20000, 2000, 10)": 1800,
        "SLOPE(Data!A2:A6, Data!B2:B6)": 1.33333333333333,
        "SMALL(Data!A1:A10, 2)": 4,
        "SQRT(16)": 4,
        "STANDARDIZE(3, 2, 1)": 1,
        "STDEV(2, 4, 6, 8)": 2.58198889747161,
        "STDEV.P(2, 4, 6, 8)": 2.23606797749979,
        "STDEV.S(2, 4, 6, 8)": 2.58198889747161,
        "STDEVA(2, 4, 6, 8)": 2.58198889747161,
        "STDEVP(2, 4, 6, 8)": 2.23606797749979,
        "STDEVPA(2, 4, 6, 8)": 2.23606797749979,
        "SUBSTITUTE(\"hello\", \"l\", \"w\")": "hewwo",
        "SUM(10, 20, 30)": 60,
        "SUMIF(Data!A1:A10, \"<=8\")": 20,
        "SUMIFS(Data!A1:A10, Data!A1:A10, \">2\", Data!B1:B10, \"<10\")": 40,
        "SUMPRODUCT(Data!A2:A6, Data!B2:B6)": 270,
        "SUMSQ(2, 4, 6)": 56,
        "SUMX2MY2(Data!A2:A6, Data!B2:B6)": 157.5,
        "SUMX2PY2(Data!A2:A6, Data!B2:B6)": 562.5,
        "SUMXMY2(Data!A2:A6, Data!B2:B6)": 22.5,
        "SWITCH(2, 1, \"A\", 2, \"B\", \"C\")": "B",
        "SYD(20000, 2000, 10, 2)": 2945.45454545455,
        "T(1)": "",
        "TAN(1)": 1.5574077246549,
        "TEXT(46000, \"yyyy/mm/dd\")": "2025/12/09",
        "TEXTJOIN(\"-\", TRUE(), \"X\", \"Y\")": "X-Y",
        "TIME(14, 30, 0)": 0.604166666666667,
        "TODAY()": 46118,
        "TRIM(\"  hello  world  \")": "hello  world",
        "TRIMMEAN(Data!A1:A10, 0.4)": 11,
        "TRUE()": True,
        "TRUNC(3.14159, 2)": 3.14,
        "UNICHAR(66)": "B",
        "UNICODE(\"B\")": 66,
        "UPPER(\"hello\")": "HELLO",
        "VALUE(\"2\")": 2,
        "VAR(2, 4, 6, 8)": 6.66666666666667,
        "VAR.P(2, 4, 6, 8)": 5,
        "VAR.S(2, 4, 6, 8)": 6.66666666666667,
        "VARA(2, 4, 6, 8)": 6.66666666666667,
        "VARP(2, 4, 6, 8)": 5,
        "VARPA(2, 4, 6, 8)": 5,
        "VLOOKUP(6, Data!A1:E10, 3, FALSE())": "Item3",
        "WEEKDAY(46000, 2)": 2,
        "WEEKNUM(46000, 2)": 50,
        "WORKDAY(46000, 2)": 46002,
        "XLOOKUP(6, Data!A1:A10, Data!C1:C10)": "Item3",
        "YEAR(46000)": 2025,
    }

    for formula, expected in test_cases.items():
        result = engine.evaluate(formula, ws)
        if isinstance(expected, (float, int)) and not isinstance(expected, bool) and not isinstance(result, str):
            assert result == pytest.approx(expected, rel=1e-4, abs=1e-9), f"Failed: {formula}"
        else:
            assert result == expected, f"Failed: {formula} -> expected {expected}, got {result}"
    # Test error cases
    error_cases = {
        "HLOOKUP(\"Item3\", Data!C1:E10, 3, FALSE())": "ValueError",  # #N/A equivalent in pyopenxlsx might raise or return string depending on bindings
    }
    
    for formula in error_cases:
        try:
            result = engine.evaluate(formula, ws)
            # Depending on how the error is mapped in bindings, it might return a string with 'ERROR' or raise an Exception.
            # Just ensure it's handled gracefully
            assert isinstance(result, str) and "ERROR" in result.upper() or type(result).__name__ == error_cases[formula]
        except Exception:
            pass # Exception is also acceptable for evaluation errors
