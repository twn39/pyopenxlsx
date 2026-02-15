import zipfile
from pyopenxlsx import Workbook, Fill


def test_fill_and_merge_cells(tmp_path):
    """
    Test setting cell fill (background color) and merging cells.
    Verifies that both styles and merge ranges are correctly preserved in the saved XLSX file.
    """
    output_file = tmp_path / "test_style_merge.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "StyleMergeTest"

    # 1. Test Merge
    ws["A1"].value = "Merged Range"
    ws.merge_cells("A1:C3")

    # 2. Test Fill (Background Color)
    # Use a specific color to verify in XML
    fill_color = "FFCC00"  # Yellow
    fill = Fill(pattern_type="solid", color=fill_color)
    style_idx = wb.add_style(fill=fill)
    ws["A1"].style_index = style_idx

    # Add another cell with a different color
    other_color = "00CCFF"  # Blue
    other_fill = Fill(pattern_type="solid", color=other_color)
    other_style_idx = wb.add_style(fill=other_fill)
    ws["E1"].value = "Single Styled"
    ws["E1"].style_index = other_style_idx

    wb.save(output_file)
    wb.close()

    # 3. Verify XML structure
    with zipfile.ZipFile(output_file, "r") as z:
        # Verify styles.xml
        styles_xml = z.read("xl/styles.xml").decode("utf-8").upper()
        assert 'PATTERNTYPE="SOLID"' in styles_xml
        # OpenXLSX normalizes to ARGB (FFFFCC00 and FF00CCFF)
        assert f'RGB="FF{fill_color}"' in styles_xml
        assert f'RGB="FF{other_color}"' in styles_xml

        # Verify sheet1.xml
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert '<mergeCell ref="A1:C3"/>' in sheet_xml
        assert 'r="A1"' in sheet_xml
        assert 'r="E1"' in sheet_xml
