import pytest
from pyopenxlsx import Workbook

def test_vml_drawing_shapes():
    wb = Workbook()
    ws = wb.active
    
    # We should have a VML drawing attached to the worksheet implicitly when we try to use shapes?
    # Wait, the VML drawing is created automatically when needed, or maybe it returns a valid XLVmlDrawing even if empty.
    vml = ws._sheet.vml_drawing()
    assert vml is not None

    # Let's add a shape
    # create_shape returns an XLShape
    shape = vml.create_shape()
    assert vml.shape_count() == 1

    # Modify properties
    shape.set_fill_color("#FF0000")
    assert shape.fill_color() == "#FF0000"

    shape.set_type("#_x0000_t202")
    assert shape.type() == "#_x0000_t202"

    shape.set_stroked(False)
    assert shape.stroked() is False

    # Client data
    client_data = shape.client_data()
    client_data.set_row(5)
    client_data.set_column(2)
    
    assert client_data.row() == 5
    assert client_data.column() == 2

    # Style
    style = shape.style()
    style.set_width(100)
    style.set_height(50)
    assert style.width() == 100
    assert style.height() == 50

    # Test reading back
    read_shape = vml.shape(0)
    assert read_shape.fill_color() == "#FF0000"

    # Delete shape
    vml.delete_shape(0)
    assert vml.shape_count() == 0
