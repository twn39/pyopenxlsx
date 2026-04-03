# Comments API

`pyopenxlsx` supports adding and interacting with standard comments, as well as the modern Threaded Comments (used in newer versions of Excel for conversational replies).

## Adding Legacy Comments

A standard, non-threaded comment attached to a cell.

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    ws.add_comment("A1", text="This cell needs review.", author="Alice")
    
    wb.save("comments.xlsx")
```

## Threaded Comments

Threaded comments support rich conversations with authors (persons), replies, and resolved states.

### Interacting with Threaded Comments

You must first register a "Person" (an author) in the workbook before adding threaded comments.

```python
from pyopenxlsx import Workbook

with Workbook() as wb:
    ws = wb.active
    
    # 1. Register a Person in the document
    persons = wb._doc.persons()
    author_id = persons.add_person("John Doe")
    
    # 2. Access the threaded comments collection for the active sheet
    threads = ws._sheet.threaded_comments()
    
    # 3. Add a top-level comment to a cell (e.g. A1)
    root_comment = threads.add_comment("A1", author_id, "Please update these figures.")
    
    # 4. Add a reply to the root comment
    # Notice we pass the parent's ID
    reply = threads.add_reply(root_comment.id, author_id, "Updated!")
    
    # 5. Mark the conversation as resolved
    root_comment.is_resolved = True
    
    wb.save("threaded_comments.xlsx")
```

### Reading Threaded Comments

You can load an existing workbook and retrieve conversation threads:

```python
from pyopenxlsx import load_workbook

with load_workbook("threaded_comments.xlsx") as wb:
    ws = wb.active
    
    threads = ws._sheet.threaded_comments()
    
    # Get the root comment on a specific cell
    c = threads.comment("A1")
    if c.valid:
        print(f"Top Comment: '{c.text}' (Resolved: {c.is_resolved})")
        
        # Get all replies
        replies = threads.replies(c.id)
        for r in replies:
            print(f"  Reply: '{r.text}'")
```

*(Note: Advanced Threaded Comments API is accessed via the internal C++ bindings `. _doc` and `._sheet` properties to offer raw maximum control over the underlying OpenXLSX objects).*