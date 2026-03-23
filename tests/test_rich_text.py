from pyopenxlsx import Workbook, XLRichText, XLRichTextRun, XLColor

def test_rich_text_basic():
    wb = Workbook()
    ws = wb.active
    
    rt = XLRichText()
    run1 = XLRichTextRun("Hello ")
    run1.bold = True
    rt.add_run(run1)
    
    run2 = XLRichTextRun("World")
    run2.italic = True
    run2.font_color = XLColor(0, 0, 255)
    rt.add_run(run2)
    
    ws.cell(1, 1).value = rt
    assert ws.cell(1, 1).value.plain_text == "Hello World"
    
    # Check runs
    runs = list(ws.cell(1, 1).value.runs)
    assert len(runs) == 2
    assert runs[0].text == "Hello "
    assert runs[0].bold is True
    assert runs[1].text == "World"
    assert runs[1].italic is True

def test_rich_text_save_load(tmp_path):
    filepath = tmp_path / "rich_text.xlsx"
    wb = Workbook()
    ws = wb.active
    
    rt = XLRichText()
    r = XLRichTextRun("Styled")
    r.bold = True
    r.font_size = 15
    r.font_name = "Arial"
    rt.add_run(r)
    rt.add_run(XLRichTextRun(" Plain"))
    
    ws.cell(1, 1).value = rt
    wb.save(filepath)
    
    wb2 = Workbook(filepath)
    val = wb2.active.cell(1, 1).value
    assert isinstance(val, XLRichText)
    assert val.plain_text == "Styled Plain"
    runs = list(val.runs)
    assert len(runs) == 2
    assert runs[0].text == "Styled"
    assert runs[0].bold is True
    assert runs[0].font_size == 15
    assert runs[0].font_name == "Arial"
    wb2.close()

def test_rich_text_clear():
    rt = XLRichText("Initial")
    assert not rt.empty()
    rt.clear()
    assert rt.empty()
    assert rt.plain_text == ""
