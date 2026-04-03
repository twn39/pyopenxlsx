import pyopenxlsx
import os


def test_add_shapes(tmp_path):
    wb = pyopenxlsx.Workbook()
    ws = wb.active

    ws.add_shape(
        row=2,
        col=2,
        shape_type="Arrow",
        name="MyArrow",
        text="Point!",
        fill_color="FF0000",
        line_width=2.5,
        rotation=90,
    )

    ws.add_shape(
        row=5,
        col=5,
        shape_type="Cloud",
        name="Cloudy",
        text="Cloud",
        width=200,
        height=150,
        flip_h=True,
    )

    file_path = tmp_path / "shapes.xlsx"
    wb.save(str(file_path))

    assert os.path.exists(file_path)
    # We can't easily parse the shape back because OpenXLSX doesn't have an iterative getter for shapes yet,
    # but creating it without exceptions and generating the file is exactly what we need to verify.
