import pyopenxlsx
import os


def test_document_context_manager(tmp_path):
    filename = tmp_path / "test_context.xlsx"

    with pyopenxlsx.XLDocument() as doc:
        doc.create(str(filename))
        doc.save()

    assert os.path.exists(filename)


def test_document_manual_close(tmp_path):
    filename = tmp_path / "test_manual_close.xlsx"

    doc = pyopenxlsx.XLDocument()
    doc.create(str(filename))
    doc.save()
    doc.close()

    assert os.path.exists(filename)


def test_document_properties(tmp_path):
    filename = tmp_path / "test_props.xlsx"
    doc = pyopenxlsx.XLDocument()
    doc.create(str(filename))

    # We don't have full property bindings yet in high-level,
    # but we can test if the basic document object works.
    assert doc.workbook() is not None
    doc.close()
