from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import (
    XLChartType,
    XLPivotTableOptions,
    XLPivotSubtotal,
    XLPivotField,
)


def test_charts(tmp_path):
    file_path = tmp_path / "test_charts.xlsx"

    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, ["Category", "Value1", "Value2"])
        ws.write_rows(2, [["A", 10, 20], ["B", 15, 25], ["C", 20, 30]])

        chart = ws._sheet.add_chart(XLChartType.Bar, "Chart1", 5, 5, 400, 300)  # type: ignore
        assert chart is not None

        chart.set_title("My Chart")
        chart.add_series_ref("Sheet1!$B$2:$B$4", "Series 1", "Sheet1!$A$2:$A$4")

        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        assert ws.cell(1, 1).value == "Category"


def test_pivot_tables(tmp_path):
    file_path = tmp_path / "test_pivot.xlsx"

    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, ["Region", "Sales"])
        ws.write_rows(
            2, [["North", 100], ["South", 200], ["North", 150], ["South", 250]]
        )

        options = XLPivotTableOptions()
        options.name = "PivotTable1"
        options.source_range = "A1:B5"
        options.target_cell = "D1"

        pf_region = XLPivotField()
        pf_region.name = "Region"

        pf_sales = XLPivotField()
        pf_sales.name = "Sales"
        pf_sales.subtotal = XLPivotSubtotal.Sum  # type: ignore

        options.rows = [pf_region]
        options.data = [pf_sales]

        ws._sheet.add_pivot_table(options)

        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        assert ws.cell(1, 1).value == "Region"
