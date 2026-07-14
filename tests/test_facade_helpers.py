"""Tests for high-level chart / CF / stream façades."""

from datetime import date, datetime

from pyopenxlsx import Workbook, add_chart, chart_type, conditional_formatting
from pyopenxlsx._openxlsx import XLChartType


def test_chart_type_string_and_enum():
    assert chart_type("bar") == XLChartType.Bar
    assert chart_type("Column") == XLChartType.Column
    assert chart_type(XLChartType.Pie) == XLChartType.Pie


def test_add_chart_string_type_and_series(tmp_path):
    path = tmp_path / "chart_facade.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.write_rows(1, [["Cat", "Val"], ["A", 1], ["B", 2]])
        chart = ws.add_chart("bar", "C1", 5, 5, 300, 200)
        chart.title("T")
        chart.series("Sheet1!$B$2:$B$3", name="S", categories_ref="Sheet1!$A$2:$A$3")
        # Module-level helper
        add_chart(
            ws,
            "line",
            "C2",
            row=20,
            col=5,
            title="Line",
            series_ref="Sheet1!$B$2:$B$3",
            cats_ref="Sheet1!$A$2:$A$3",
        )
        wb.save(path)
    assert path.exists()


def test_conditional_formatting_builders(tmp_path):
    path = tmp_path / "cf_facade.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, [10, 20, 30])
        ws.add_conditional_formatting(
            "A1:C1", conditional_formatting.color_scale("#FF0000", (0, 255, 0))
        )
        ws.add_conditional_formatting(
            "A1:C1", conditional_formatting.data_bar((0, 0, 255))
        )
        ws.add_conditional_formatting(
            "A1:C1", conditional_formatting.cell_is(">", "15")
        )
        ws.add_conditional_formatting(
            "A1:C1", conditional_formatting.formula_rule("A1>10")
        )
        wb.save(path)
    with Workbook(path) as wb:
        assert wb.active.cell(1, 1).value == 10


def test_stream_writer_date_coercion(tmp_path):
    path = tmp_path / "stream_dates.xlsx"
    d = date(2023, 5, 1)
    dt = datetime(2023, 5, 1, 12, 0, 0)
    with Workbook() as wb:
        ws = wb.active
        with ws.stream_writer() as writer:
            writer.append_row(["label", d, dt])
        wb.save(path)

    with Workbook(path) as wb:
        ws = wb.active
        # Stream writes go to sheet XML; after reload, date formats should apply
        # for auto_date_formats path (style tuples on write).
        assert ws.cell(1, 1).value == "label"
        # Serials present
        assert isinstance(ws.get_cell_value(1, 2), (int, float))
        assert isinstance(ws.get_cell_value(1, 3), (int, float))
