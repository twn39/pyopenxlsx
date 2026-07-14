"""Tests for pivot builder, Chart façade, and style registry extraction."""

from pyopenxlsx import (
    Chart,
    PivotTableBuilder,
    Workbook,
    add_chart,
    pivot_subtotal,
    pivot_table,
)
from pyopenxlsx._openxlsx import XLPivotSubtotal
from pyopenxlsx._style_registry import register_cell_style
from pyopenxlsx.styles import Font


def test_pivot_subtotal_resolve():
    assert pivot_subtotal("sum") == XLPivotSubtotal.Sum
    assert pivot_subtotal("COUNT") == XLPivotSubtotal.Count
    assert pivot_subtotal(XLPivotSubtotal.Average) == XLPivotSubtotal.Average


def test_pivot_builder_chain(tmp_path):
    path = tmp_path / "pivot_builder.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, ["Region", "Sales"])
        ws.write_rows(
            2, [["North", 100], ["South", 200], ["North", 150], ["South", 250]]
        )
        PivotTableBuilder("PivotTable1", "Sheet1!A1:B5", "D1").rows(
            "Region"
        ).data("Sales", subtotal="sum", num_fmt_id=4).style(
            "PivotStyleMedium9"
        ).data_on_rows(True).show_headers(rows=True).stripes(
            columns=True
        ).compact(False).add_to(ws)
        wb.save(path)

    with Workbook(path) as wb:
        assert wb.active.cell(1, 1).value == "Region"


def test_pivot_table_helper_and_add_options(tmp_path):
    path = tmp_path / "pivot_helper.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, ["Region", "Sales"])
        ws.write_rows(2, [["North", 100], ["South", 200]])
        builder = pivot_table(
            "P2",
            "Sheet1!A1:B3",
            "E1",
            rows="Region",
            data="Sales",
            style="PivotStyleMedium2",
        )
        ws.add_pivot_table(builder)
        wb.save(path)
    assert path.exists()


def test_chart_fluent_api(tmp_path):
    path = tmp_path / "chart_fluent.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.write_rows(1, [["Cat", "V1", "V2"], ["A", 10, 20], ["B", 15, 25]])
        chart = (
            Chart(ws.add_chart("bar", "C1", 5, 5, 400, 300, wrap=False))
            .title("Sales")
            .legend("bottom")
            .series("Sheet1!$B$2:$B$3", name="V1", categories_ref="Sheet1!$A$2:$A$3")
            .series("Sheet1!$C$2:$C$3", name="V2", categories_ref="Sheet1!$A$2:$A$3")
            .overlap(10)
            .data_labels(value=True)
        )
        assert isinstance(chart, Chart)
        # ws.add_chart wraps by default
        wrapped = ws.add_chart("column", "C2", 20, 5)
        assert isinstance(wrapped, Chart)
        wrapped.series_many(
            [
                ("Sheet1!$B$2:$B$3", "V1"),
                {"values": "Sheet1!$C$2:$C$3", "name": "V2"},
            ],
            categories_ref="Sheet1!$A$2:$A$3",
        )
        add_chart(
            ws,
            "line",
            "C3",
            row=20,
            col=15,
            title="Line",
            series=[
                ("Sheet1!$B$2:$B$3", "V1"),
            ],
            cats_ref="Sheet1!$A$2:$A$3",
            legend="right",
        )
        wb.save(path)
    assert path.exists()


def test_style_registry_matches_workbook_add_style():
    with Workbook() as wb:
        font = Font(name="Arial", size=14, bold=True)
        idx1 = wb.add_style(font=font)
        font2 = Font(name="Arial", size=14, bold=True)
        idx2 = register_cell_style(wb.styles, font=font2)
        assert isinstance(idx1, int)
        assert isinstance(idx2, int)
        assert idx1 >= 0
        assert idx2 >= 0
        # Number format path still works
        idx3 = wb.add_style(number_format="0.00%")
        assert idx3 >= 0
