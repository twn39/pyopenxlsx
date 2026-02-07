import pyopenxlsx
import os


def test_content_types(tmp_path):
    filename = tmp_path / "test_ct.xlsx"
    doc = pyopenxlsx.XLDocument()
    doc.create(str(filename))

    ct = doc.content_types()
    items = ct.get_content_items()

    # Check default items
    paths = [item.path() for item in items]
    assert "/xl/workbook.xml" in paths
    assert "/xl/styles.xml" in paths
    assert "/xl/theme/theme1.xml" in paths

    # Add a custom override
    ct.add_override("/xl/custom.xml", pyopenxlsx.XLContentType.Worksheet)

    items2 = ct.get_content_items()
    paths2 = [item.path() for item in items2]
    assert "/xl/custom.xml" in paths2

    # Find specific item
    item = ct.content_item("/xl/custom.xml")
    assert item.path() == "/xl/custom.xml"
    assert item.type() == pyopenxlsx.XLContentType.Worksheet

    # Delete override
    ct.delete_override("/xl/custom.xml")
    items3 = ct.get_content_items()
    paths3 = [item.path() for item in items3]
    assert "/xl/custom.xml" not in paths3

    # Test multiple overrides
    ct.add_override("/xl/custom1.xml", pyopenxlsx.XLContentType.Worksheet)
    ct.add_override("/xl/custom2.xml", pyopenxlsx.XLContentType.Chartsheet)

    items4 = ct.get_content_items()
    paths4 = [item.path() for item in items4]
    assert "/xl/custom1.xml" in paths4
    assert "/xl/custom2.xml" in paths4

    # Delete by item
    item_to_delete = ct.content_item("/xl/custom1.xml")
    ct.delete_override(item_to_delete)

    items5 = ct.get_content_items()
    paths5 = [item.path() for item in items5]
    assert "/xl/custom1.xml" not in paths5
    assert "/xl/custom2.xml" in paths5

    doc.save()
    doc.close()


def test_content_type_enum():
    assert pyopenxlsx.XLContentType.Workbook is not None
    assert pyopenxlsx.XLContentType.Worksheet is not None
    assert pyopenxlsx.XLContentType.Styles is not None
    assert pyopenxlsx.XLContentType.SharedStrings is not None
