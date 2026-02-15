import zipfile
from pyopenxlsx import (
    Workbook,
    Font,
    Fill,
    Alignment,
    Border,
    Side,
    XLPatternType,
    XLAlignmentStyle,
    XLLineStyle,
)


def test_font_properties():
    """Test Font class property accessors and defaults."""
    f = Font(name="Calibri", size=12, bold=True, italic=False, color="FF0000")
    assert f.name() == "Calibri"
    assert f.size() == 12
    assert f.bold() is True
    assert f.italic() is False
    assert f.color().hex().lower() == "ffff0000"  # XLColor normalizes to ARGB

    f.set_bold(False)
    assert f.bold() is False


def test_fill_properties():
    """Test Fill class property accessors and enum mapping."""
    fill = Fill(pattern_type="solid", color="00FF00", background_color="000000")
    assert fill.pattern_type() == XLPatternType.Solid
    assert fill.color().hex().lower() == "ff00ff00"

    fill.set_pattern_type(XLPatternType.MediumGray)
    assert fill.pattern_type() == XLPatternType.MediumGray


def test_alignment_properties():
    """Test Alignment class property accessors and enum mapping."""
    align = Alignment(horizontal="center", vertical="top", wrap_text=True)
    assert align.horizontal() == XLAlignmentStyle.Center
    assert align.vertical() == XLAlignmentStyle.Top
    assert align.wrap_text() is True

    align.set_horizontal(XLAlignmentStyle.Left)
    assert align.horizontal() == XLAlignmentStyle.Left


def test_border_properties():
    """Test Border and Side class properties."""
    side = Side(style="thin", color="FF0000")
    assert side.style() == XLLineStyle.Thin
    assert side.color().hex().lower() == "ffff0000"

    border = Border(left=side, right=side, top="thick", bottom="dashed")
    assert border.left().style() == XLLineStyle.Thin
    assert border.right().style() == XLLineStyle.Thin
    assert border.top().style() == XLLineStyle.Thick
    assert border.bottom().style() == XLLineStyle.Dashed


def test_workbook_add_style_font(tmp_path):
    """Test adding a font style to a workbook and verifying XML."""
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Styled"

    font = Font(name="Arial", size=14, bold=True)
    style_idx = wb.add_style(font=font)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_font.xlsx"
    wb.save(output)
    wb.close()

    # Verify XML content
    with zipfile.ZipFile(output, "r") as z:
        styles_xml = z.read("xl/styles.xml").decode("utf-8")
        assert 'val="Arial"' in styles_xml
        assert 'val="14"' in styles_xml
        # Bold might be <b/> or <b val="1"/> or <b val="true"/>
        assert "<b" in styles_xml


def test_workbook_add_style_border_no_empty_style(tmp_path):
    """
    Regression test: Ensure that borders with style=None do not produce style="" attribute in XML.
    Using 'thick' which maps to a known style, and 'none' which should be ignored or handled.
    """
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Border"

    # Create a border with explicit None/empty sides
    border = Border(left=Side(style="thick", color="000000"))
    # Right, Top, Bottom are None by default

    style_idx = wb.add_style(border=border)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_border_xml.xlsx"
    wb.save(output)
    wb.close()

    with zipfile.ZipFile(output, "r") as z:
        styles_xml = z.read("xl/styles.xml").decode("utf-8")
        # Ensure we don't see style=""
        assert 'style=""' not in styles_xml
        # Ensure we see the thick border
        assert 'style="thick"' in styles_xml


def test_workbook_add_style_fill_pattern(tmp_path):
    """Test adding a fill style and verifying XML."""
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Filled"

    fill = Fill(pattern_type="solid", color="FFFF00")
    style_idx = wb.add_style(fill=fill)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_fill.xlsx"
    wb.save(output)
    wb.close()

    with zipfile.ZipFile(output, "r") as z:
        styles_xml = z.read("xl/styles.xml").decode("utf-8")
        assert 'patternType="solid"' in styles_xml
        # Hex color might be FFFFFF00 or ffffff00
        assert 'rgb="FFFFFF00"' in styles_xml or 'rgb="ffffff00"' in styles_xml


def test_workbook_add_style_alignment(tmp_path):
    """Test adding alignment style."""
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Aligned"

    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    style_idx = wb.add_style(alignment=align)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_align.xlsx"
    wb.save(output)
    wb.close()

    with zipfile.ZipFile(output, "r") as z:
        styles_xml = z.read("xl/styles.xml").decode("utf-8")
        assert 'horizontal="center"' in styles_xml
        assert 'vertical="center"' in styles_xml
        # wrapText might be 1 or true
        assert 'wrapText="1"' in styles_xml or 'wrapText="true"' in styles_xml


def test_complex_style_object(tmp_path):
    """Test applying a Style object containing logical groups."""
    from pyopenxlsx.styles import Style

    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Complex"

    font = Font(bold=True)
    fill = Fill(pattern_type="solid", color="EEEEEE")
    align = Alignment(horizontal="right")

    style = Style(font=font, fill=fill, alignment=align)
    style_idx = wb.add_style(
        font=style
    )  # passing Style object as first arg (param name mismatch handled in logic)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_complex.xlsx"
    wb.save(output)
    wb.close()

    with zipfile.ZipFile(output, "r") as z:
        styles_xml = z.read("xl/styles.xml").decode("utf-8")
        assert "<b" in styles_xml
        assert 'patternType="solid"' in styles_xml
        assert 'horizontal="right"' in styles_xml
