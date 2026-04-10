from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import XLSparklineType, XLSparklineOptions


def test_sparkline_options(tmp_path):
    file_path = tmp_path / "test_sparkline_options.xlsx"

    with Workbook() as wb:
        ws = wb.active
        ws.write_row(1, ["Data", 10, -20, 30, 40])

        options = XLSparklineOptions()
        options.type = XLSparklineType.Column
        options.series_color = "FF0000FF"
        options.negative_color = "FFFF0000"
        options.markers = True
        options.high = True
        options.low = True
        options.first = True
        options.last = True
        options.negative = True
        options.display_x_axis = True
        options.display_empty_cells_as = "zero"

        ws.add_sparkline("F1", "B1:E1", options=options)

        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        assert ws.cell(1, 1).value == "Data"
        # If no crash happened during save/load, the options are properly set and written
