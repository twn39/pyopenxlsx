from pyopenxlsx import (
    Workbook,
    Font,
    Fill,
    Alignment,
    XLPatternType,
    XLAlignmentStyle,
)


def test_font_style(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Hello"

    font = Font(name="Courier New", size=14, bold=True, italic=True, color="FF0000")
    style_idx = wb.add_style(font=font)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_font.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_fill_style(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Fill"

    fill = Fill(pattern_type=XLPatternType.Solid, color="FFFF00")
    style_idx = wb.add_style(fill=fill)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_fill.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_alignment_style(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Align"

    align = Alignment(
        horizontal=XLAlignmentStyle.Center,
        vertical=XLAlignmentStyle.Center,
        wrap_text=True,
    )
    style_idx = wb.add_style(alignment=align)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_align.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_cell_merging(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["B2"].value = "Merged"
    ws.merge_cells("B2:D4")

    # We don't have a direct "is_merged" API yet in high-level,
    # but we can verify it doesn't crash and saves correctly.
    output = tmp_path / "test_merge.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_column_row_format(tmp_path):
    wb = Workbook()
    ws = wb.active

    font = Font(bold=True)
    style_idx = wb.add_style(font=font)

    ws.set_column_format("C", style_idx)
    ws.set_row_format(5, style_idx)

    output = tmp_path / "test_format.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_range_styling(tmp_path):
    wb = Workbook()
    ws = wb.active

    style_idx = wb.add_style(font=Font(italic=True))
    r = ws.range("A1", "B2")
    for cell in r:
        cell.value = "Test"
        cell.style_index = style_idx

    output = tmp_path / "test_range.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_border_style(tmp_path):
    from pyopenxlsx import Border, Side, XLLineStyle

    wb = Workbook()
    ws = wb.active
    ws["B2"].value = "Border Check"

    thin = Side(style=XLLineStyle.Thin, color="FF0000")
    double = Side(style=XLLineStyle.Double, color="0000FF")

    border = Border(left=thin, right=double, top="thick", bottom="dashed")
    style_idx = wb.add_style(border=border)
    ws["B2"].style_index = style_idx

    output = tmp_path / "test_border.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_number_format_style(tmp_path):
    wb = Workbook()
    ws = wb.active

    # Standard Date Format
    ws["A1"].value = 45000  # Some date logic usually implies internal float/int
    style_date = wb.add_style(number_format="yyyy-mm-dd")
    ws["A1"].style_index = style_date

    # Custom Currency
    ws["A2"].value = 1234.5678
    style_currency = wb.add_style(number_format="#,##0.00 $")
    ws["A2"].style_index = style_currency

    output = tmp_path / "test_numfmt.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_styles_edge_cases():
    from pyopenxlsx.styles import Font, Fill, Side, Border, is_date_format
    from pyopenxlsx._openxlsx import XLColor

    # Font with XLColor object
    color = XLColor("FF0000")
    Font(color=color)

    # Fill with XLColor and background color
    Fill(color=color, background_color="00FF00")
    Fill(background_color=XLColor("0000FF"))

    # is_date_format edge cases
    assert is_date_format(None) is False
    assert is_date_format(1.5) is False

    # Side with XLColor object
    Side(color=color)
    Side(color=None)  # Should default to black

    # Border with direct style strings
    Border(
        left="thin", right="thick", top="dashed", bottom="double", diagonal="medium"
    )
