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
    color = fill.color()
    assert color is not None and color.hex().lower() == "ff00ff00"
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


def test_font_underline_strikethrough():
    from pyopenxlsx import XLUnderlineStyle

    f = Font(underline=True, strikethrough=True)
    assert f.underline() == XLUnderlineStyle.Single
    assert f.strikethrough() is True
    f.set_underline("double")
    assert f.underline() == XLUnderlineStyle.Double
    f.set_underline(None)
    assert f.underline() == getattr(XLUnderlineStyle, "None")


def test_alignment_indent_rotation_shrink():
    align = Alignment(
        horizontal="left",
        indent=2,
        text_rotation=45,
        shrink_to_fit=True,
    )
    assert align.indent() == 2
    assert align.text_rotation() == 45
    assert align.shrink_to_fit() is True
    align.set_indent(3)
    align.set_text_rotation(90)
    align.set_shrink_to_fit(False)
    assert align.indent() == 3
    assert align.text_rotation() == 90
    assert align.shrink_to_fit() is False


def test_workbook_add_style_font_extras(tmp_path):
    """Underline/strikethrough and alignment extras land in styles.xml."""
    from pyopenxlsx import XLUnderlineStyle

    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Styled"

    font = Font(
        name="Arial",
        size=12,
        bold=True,
        underline=XLUnderlineStyle.Single,
        strikethrough=True,
    )
    align = Alignment(indent=1, text_rotation=15, shrink_to_fit=True, wrap_text=True)
    style_idx = wb.add_style(font=font, alignment=align)
    ws["A1"].style_index = style_idx

    output = tmp_path / "test_font_extras.xlsx"
    wb.save(output)
    wb.close()

    with zipfile.ZipFile(output, "r") as z:
        styles_xml = z.read("xl/styles.xml").decode("utf-8")
        assert (
            "strike" in styles_xml
            or 'strike="' in styles_xml
            or "<strike" in styles_xml
        )
        # underline single is typically u="single" or <u val="single"/>
        assert "underline" in styles_xml or "u=" in styles_xml or "<u" in styles_xml
        assert (
            "indent" in styles_xml
            or "textRotation" in styles_xml
            or "shrinkToFit" in styles_xml
        )


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


def test_compact_styles(tmp_path):
    """Test deduplicating and compacting cell formats."""
    wb = Workbook()
    ws = wb.active

    # Create a shared font and two identical cell formats referencing it
    f_idx = wb.styles.fonts().create()
    font = wb.styles.fonts().font_by_index(f_idx)
    font.set_bold(True)
    font.set_size(14)

    s1 = wb.add_style(font=f_idx)
    s2 = wb.add_style(font=f_idx)
    # Add an unused orphan style
    _ = wb.add_style(font=Font(italic=True, size=20))

    ws["A1"].value = "Header 1"
    ws["A1"].style_index = s1

    ws["B1"].value = "Header 2"
    ws["B1"].style_index = s2

    # Before compaction: 4 formats (0, s1, s2, orphan)
    assert wb.styles.cell_formats().count() == 4

    # Verify compaction cleans up orphans and remaps duplicate XFs
    wb.compact_styles()

    # After compaction: 2 formats (0, unified s1/s2)
    assert wb.styles.cell_formats().count() == 2

    output = tmp_path / "test_compact.xlsx"
    wb.save(output, compact=True)
    wb.close()

    with Workbook(output) as loaded_wb:
        loaded_ws = loaded_wb.active
        assert loaded_ws["A1"].value == "Header 1"
        assert loaded_ws["B1"].value == "Header 2"
        # Style indices should be valid, deduplicated, and identical
        assert loaded_ws["A1"].style_index == loaded_ws["B1"].style_index == 1
        assert loaded_wb.styles.cell_formats().count() == 2
