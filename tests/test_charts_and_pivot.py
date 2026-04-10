from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import (
    XLChartType,
    XLPivotTableOptions,
    XLPivotSubtotal,
    XLPivotField,
    XLAxisOrientation,
    XLAxisCrosses,
)


def test_charts(tmp_path):
    file_path = tmp_path / "test_charts.xlsx"

    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, ["Category", "Value1", "Value2", "Size"])
        ws.write_rows(2, [["A", 10, 20, 5], ["B", 15, 25, 10], ["C", 20, 30, 15]])

        # Test Bar chart with advanced properties
        chart = ws._sheet.add_chart(XLChartType.Bar, "Chart1", 5, 5, 400, 300)
        assert chart is not None

        chart.set_title("My Chart")
        chart.add_series_ref("Sheet1!$B$2:$B$4", "Series 1", "Sheet1!$A$2:$A$4")
        
        # New chart methods
        chart.set_overlap(20)
        
        # Test axis settings
        x_axis = chart.x_axis()
        x_axis.set_major_unit(5)
        x_axis.set_minor_unit(1)
        x_axis.set_log_scale(0)
        x_axis.set_date_axis(False)
        x_axis.set_orientation(XLAxisOrientation.MaxMin)
        x_axis.set_crosses(XLAxisCrosses.AutoZero)
        x_axis.set_number_format("0.00", False)
        
        # Test 3D and Doughnut chart methods
        doughnut_chart = ws._sheet.add_chart(XLChartType.Doughnut, "Chart2", 5, 15, 400, 300)
        doughnut_chart.set_hole_size(50)
        
        chart3d = ws._sheet.add_chart(XLChartType.Bar3D, "Chart3", 20, 5, 400, 300)
        chart3d.set_rotation(45, 120, 30)

        # Test Bubble series
        bubble_chart = ws._sheet.add_chart(XLChartType.Bubble, "Chart4", 20, 15, 400, 300)
        bubble_chart.add_bubble_series("Sheet1!$B$2:$B$4", "Sheet1!$C$2:$C$4", "Sheet1!$D$2:$D$4", "Bubble Series")

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
        
        # New options
        options.data_on_rows = True
        options.show_row_headers = True
        options.show_col_stripes = True
        options.pivot_table_style_name = "PivotStyleMedium9"
        options.compact_data = False

        pf_region = XLPivotField()
        pf_region.name = "Region"

        pf_sales = XLPivotField()
        pf_sales.name = "Sales"
        pf_sales.subtotal = XLPivotSubtotal.Sum  # type: ignore
        pf_sales.num_fmt_id = 4  # e.g., '#,##0.00'

        options.rows = [pf_region]
        options.data = [pf_sales]

        ws._sheet.add_pivot_table(options)

        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        assert ws.cell(1, 1).value == "Region"

def test_stock_and_surface_charts(tmp_path):
    """Test newly fixed StockOHLC and Surface3D charts in OpenXLSX."""
    file_path = tmp_path / "test_stock_surface.xlsx"

    with Workbook() as wb:
        ws = wb.active
        
        # --- Stock Data ---
        ws.write_row(1, ["Date", "Open", "High", "Low", "Close"])
        ws.write_rows(2, [
            ["01/01", 100, 110, 90, 105],
            ["01/02", 105, 115, 100, 110],
            ["01/03", 110, 120, 105, 115]
        ])

        stock_chart = ws._sheet.add_chart(XLChartType.StockOHLC, "Stock Performance", 2, 7, 500, 350)
        stock_chart.add_series_ref("Sheet1!$B$2:$B$4", "Open", "Sheet1!$A$2:$A$4")
        stock_chart.add_series_ref("Sheet1!$C$2:$C$4", "High", "Sheet1!$A$2:$A$4")
        stock_chart.add_series_ref("Sheet1!$D$2:$D$4", "Low", "Sheet1!$A$2:$A$4")
        stock_chart.add_series_ref("Sheet1!$E$2:$E$4", "Close", "Sheet1!$A$2:$A$4")

        # --- Surface Data ---
        ws.write_row(10, ["", "Col1", "Col2", "Col3"])
        ws.write_rows(11, [
            ["Row1", 1, 2, 3],
            ["Row2", 4, 5, 6],
            ["Row3", 7, 8, 9]
        ])

        surf_chart = ws._sheet.add_chart(XLChartType.Surface3D, "Surface 3D", 10, 7, 500, 350)
        surf_chart.add_series_ref("Sheet1!$B$11:$B$13", "Sheet1!$B$10", "Sheet1!$A$11:$A$13")
        surf_chart.add_series_ref("Sheet1!$C$11:$C$13", "Sheet1!$C$10", "Sheet1!$A$11:$A$13")
        surf_chart.add_series_ref("Sheet1!$D$11:$D$13", "Sheet1!$D$10", "Sheet1!$A$11:$A$13")

        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        assert ws.cell(1, 1).value == "Date"
        assert ws.cell(11, 2).value == 1
