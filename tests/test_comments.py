import pyopenxlsx
import os


def test_comments(tmp_path):
    filename = tmp_path / "test_comments.xlsx"
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

    wb.save(str(filename))
    wb.close()

    # Re-open and check
    wb2 = pyopenxlsx.load_workbook(str(filename))
    ws2 = wb2.active
    assert ws2._sheet.comments().count() == 1
    assert ws2["A1"].comment == "This is a comment"
    wb2.close()


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

def test_comment_anchor_and_size(tmp_path):
    import zipfile
    filename = tmp_path / "test_anchor.xlsx"
    wb = pyopenxlsx.Workbook()
    ws = wb.active
    
    ws["B2"].comment = "Large Comment"
    shape = ws._sheet.comments().shape("B2")
    
    # Set a custom anchor
    anchor_str = "2, 10, 2, 5, 8, 10, 10, 5"
    shape.client_data().set_anchor(anchor_str)
    
    # Set custom size (even if Anchor takes priority, we test the API)
    shape.style().set_width(300)
    shape.style().set_height(150)
    
    wb.save(str(filename))
    wb.close()
    
    # Verify XML content
    with zipfile.ZipFile(filename, "r") as z:
        vml = z.read("xl/drawings/vmlDrawing1.vml").decode("utf-8")
        assert f"<x:Anchor>{anchor_str}</x:Anchor>" in vml
        # C++ set_width/set_height adds 'pt' and might have the double semicolon quirk,
        # but we check if our values are present
        assert "width:300pt" in vml or "300" in vml
        assert "height:150pt" in vml or "150" in vml

