import pyopenxlsx


def test_threaded_comments(tmp_path):
    wb = pyopenxlsx.Workbook()
    ws = wb.active

    persons = wb._doc.persons()
    person_id = persons.add_person("John Doe")
    assert person_id != ""
    assert persons.person(person_id).display_name == "John Doe"

    threads = ws._sheet.threaded_comments()

    # Add top level comment
    c1 = threads.add_comment("A1", person_id, "This is a thread!")
    assert c1.valid
    assert c1.text == "This is a thread!"
    assert c1.ref == "A1"

    # Add reply
    c2 = threads.add_reply(c1.id, person_id, "This is a reply")
    assert c2.valid
    assert c2.parent_id == c1.id
    assert c2.text == "This is a reply"

    # Test retrieving
    c1_ret = threads.comment("A1")
    assert c1_ret.id == c1.id

    replies = threads.replies(c1.id)
    assert len(replies) == 1
    assert replies[0].id == c2.id

    # Test resolved
    assert not c1.is_resolved
    c1.is_resolved = True
    assert c1.is_resolved

    # Save and reload
    fn = tmp_path / "test_threads.xlsx"
    wb.save(str(fn))

    wb2 = pyopenxlsx.load_workbook(str(fn))
    ws2 = wb2.active

    threads2 = ws2._sheet.threaded_comments()
    c1_loaded = threads2.comment("A1")
    assert c1_loaded.valid
    assert c1_loaded.text == "This is a thread!"
    assert c1_loaded.is_resolved

    replies2 = threads2.replies(c1_loaded.id)
    assert len(replies2) == 1
    assert replies2[0].text == "This is a reply"
