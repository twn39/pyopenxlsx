import pyopenxlsx
import os


def test_comments():
    wb = pyopenxlsx.Workbook()
    ws = wb.active

    # Add author
    author_idx = ws._sheet.comments().add_author("Antigravity")
    assert author_idx == 0
    assert ws._sheet.comments().author_count() == 1
    assert ws._sheet.comments().author(0) == "Antigravity"

    # Set comment via XLComments.set
    ws._sheet.comments().set("A1", "This is a comment", author_id=0)
    assert ws._sheet.comments().count() == 1
    assert ws._sheet.comments().get("A1") == "This is a comment"

    # Get comment via XLComment class
    comment = ws._sheet.comments().get(0)
    assert comment.valid()
    assert comment.ref() == "A1"
    assert comment.text() == "This is a comment"
    assert comment.author_id() == 0

    # Convenience property on Cell
    cell = ws.cell(2, 2)  # B2
    cell.comment = "Another comment"
    assert cell.comment == "Another comment"
    assert ws._sheet.comments().count() == 2

    # Overwrite comment
    cell.comment = "Updated comment"
    assert cell.comment == "Updated comment"
    assert ws._sheet.comments().count() == 2

    # Delete comment
    cell.comment = None
    assert cell.comment is None
    assert ws._sheet.comments().count() == 1

    wb.save("test_comments.xlsx")

    # Re-open and check
    wb2 = pyopenxlsx.load_workbook("test_comments.xlsx")
    ws2 = wb2.active
    assert ws2._sheet.comments().count() == 1
    assert ws2["A1"].comment == "This is a comment"

    if os.path.exists("test_comments.xlsx"):
        os.remove("test_comments.xlsx")


def test_comments_overloads():
    wb = pyopenxlsx.Workbook()
    ws = wb.active

    comments = ws._sheet.comments()
    comments.add_author("Author1")
    comments.set("A1", "Comment 1", author_id=0)
    comments.set("B2", "Comment 2", author_id=0)

    # Test get by index
    c0 = comments.get(0)
    assert c0.text() == "Comment 1"

    # Test get by reference (returns string in C++)
    ca1 = comments.get("A1")
    assert ca1 == "Comment 1"

    cb2 = comments.get("B2")
    assert cb2 == "Comment 2"

    # Test count
    assert comments.count() == 2
