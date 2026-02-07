from pyopenxlsx import Workbook, XLProperty


def test_properties_basic():
    wb = Workbook()
    props = wb.properties

    props.title = "Test Title"
    props.creator = "Test Creator"
    props.subject = "Test Subject"
    props.description = "Test Description"
    props.keywords = "Test Keywords"
    props.category = "Test Category"
    props.company = "Test Company"

    assert props.title == "Test Title"
    assert props.creator == "Test Creator"
    assert props.subject == "Test Subject"
    assert props.description == "Test Description"
    assert props.keywords == "Test Keywords"
    assert props.category == "Test Category"
    assert props.company == "Test Company"


def test_properties_save_load(tmp_path):
    path = tmp_path / "test_props.xlsx"

    wb = Workbook()
    wb.properties.title = "Saved Title"
    wb.properties.creator = "Saved Creator"
    wb.save(path)

    wb2 = Workbook(path)
    assert wb2.properties.title == "Saved Title"
    assert wb2.properties.creator == "Saved Creator"


def test_low_level_properties():
    wb = Workbook()
    doc = wb._doc

    doc.set_property(XLProperty.Title, "Low Level Title")
    assert doc.property(XLProperty.Title) == "Low Level Title"

    doc.delete_property(XLProperty.Title)
    # OpenXLSX property() returns empty string if not found or creates it?
    # In XLProperties::property, it appends if empty.
    assert doc.property(XLProperty.Title) == ""


def test_app_properties():
    wb = Workbook()
    app_props = wb._doc.app_properties()

    app_props.set_property("Company", "My Company")
    assert app_props.property("Company") == "My Company"

    app_props.set_property("Manager", "My Manager")
    assert app_props.property("Manager") == "My Manager"

    app_props.delete_property("Manager")
    assert app_props.property("Manager") == ""


def test_core_properties():
    wb = Workbook()
    core_props = wb._doc.core_properties()

    core_props.set_property("dc:title", "Core Title")
    assert core_props.property("dc:title") == "Core Title"

    core_props.delete_property("dc:title")
    assert core_props.property("dc:title") == ""


def test_properties_interface_and_caching(tmp_path):
    from pyopenxlsx import load_workbook

    path = tmp_path / "test_interface.xlsx"
    wb = Workbook()

    # Test caching
    props1 = wb.properties
    props2 = wb.properties
    assert props1 is props2

    # Test dictionary interface with strings
    wb.properties["Title"] = "Dict Title"
    assert wb.properties["Title"] == "Dict Title"
    assert wb.properties.title == "Dict Title"

    # Test case insensitivity and underscores
    wb.properties["Last_Modified_By"] = "Tester"
    assert wb.properties["last modified by"] == "Tester"

    # Test custom app properties (string based)
    wb.properties["CustomProp"] = "CustomValue"
    assert wb.properties["CustomProp"] == "CustomValue"

    wb.save(path)

    wb2 = load_workbook(path)
    assert wb2.properties["Title"] == "Dict Title"
    assert wb2.properties["CustomProp"] == "CustomValue"


def test_properties_deletion():
    wb = Workbook()
    wb.properties.company = "To Be Deleted"
    assert wb.properties.company == "To Be Deleted"

    del wb.properties["Company"]
    assert wb.properties.company == ""
