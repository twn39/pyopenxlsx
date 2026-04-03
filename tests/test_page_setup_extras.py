from pyopenxlsx import Workbook, XLPageOrientation


def test_page_setup_extras():
    wb = Workbook()
    ws = wb.active

    # Margins
    m = ws.page_margins
    m.left = 1.0
    m.right = 2.0
    m.top = 3.0
    m.bottom = 4.0
    m.header = 5.0
    m.footer = 6.0
    assert m.left == 1.0
    assert m.right == 2.0
    assert m.top == 3.0
    assert m.bottom == 4.0
    assert m.header == 5.0
    assert m.footer == 6.0

    # Print options
    po = ws.print_options
    po.grid_lines = True
    assert po.grid_lines
    po.headings = True
    assert po.headings
    po.horizontal_centered = True
    assert po.horizontal_centered
    po.vertical_centered = True
    assert po.vertical_centered

    # Page setup
    ps = ws.page_setup
    ps.paper_size = 9  # A4
    assert ps.paper_size == 9

    ps.orientation = XLPageOrientation.Landscape
    assert ps.orientation == XLPageOrientation.Landscape

    ps.scale = 150
    assert ps.scale == 150

    ps.fit_to_width = 2
    assert ps.fit_to_width == 2

    ps.fit_to_height = 3
    assert ps.fit_to_height == 3

    ps.black_and_white = True
    assert ps.black_and_white
