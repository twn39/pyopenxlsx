import pyopenxlsx
import os


def test_document_context_manager():
    filename = "test_context.xlsx"
    if os.path.exists(filename):
        os.remove(filename)

    with pyopenxlsx.XLDocument() as doc:
        doc.create(filename)
        doc.save()

    assert os.path.exists(filename)
    os.remove(filename)


def test_document_manual_close():
    filename = "test_manual_close.xlsx"
    if os.path.exists(filename):
        os.remove(filename)

    doc = pyopenxlsx.XLDocument()
    doc.create(filename)
    doc.save()
    doc.close()

    assert os.path.exists(filename)
    os.remove(filename)


def test_document_properties():
    doc = pyopenxlsx.XLDocument()
    doc.create("test_props.xlsx")

    # We don't have full property bindings yet in high-level,
    # but we can test if the basic document object works.
    assert doc.workbook() is not None
    doc.close()

    if os.path.exists("test_props.xlsx"):
        os.remove("test_props.xlsx")
