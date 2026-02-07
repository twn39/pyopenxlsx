from pyopenxlsx import Workbook, load_workbook, Protection


def test_sheet_protection_basic(tmp_path):
    """Test basic sheet protection without password."""
    filename = tmp_path / "test_prot.xlsx"
    wb = Workbook()
    ws = wb.active

    ws.protect()
    assert ws.protection["protected"] is True

    wb.save(filename)

    wb2 = load_workbook(str(filename))
    assert wb2.active.protection["protected"] is True


def test_sheet_protection_password(tmp_path):
    """Test sheet protection with password."""
    filename = tmp_path / "test_pw.xlsx"
    wb = Workbook()
    ws = wb.active

    ws.protect(password="secret")
    assert ws.protection["password_set"] is True

    wb.save(filename)

    wb2 = load_workbook(str(filename))
    assert wb2.active.protection["password_set"] is True

    wb2.active.unprotect()
    assert wb2.active.protection["protected"] is False
    assert wb2.active.protection["password_set"] is False


def test_cell_locking_style(tmp_path):
    """Test applying cell locking via styles."""
    wb = Workbook()
    ws = wb.active

    # Create an unlocked style
    unlocked_idx = wb.add_style(protection=Protection(locked=False))

    ws.cell(1, 1).value = "Unlocked"
    ws.cell(1, 1).style_index = unlocked_idx

    # Create a hidden style
    hidden_idx = wb.add_style(protection=Protection(hidden=True))
    ws.cell(1, 2).value = "Hidden"
    ws.cell(1, 2).style_index = hidden_idx

    filename = tmp_path / "test_cell_prot.xlsx"
    wb.save(filename)

    # Reload and verify style indices (basic check)
    wb2 = load_workbook(str(filename))
    ws2 = wb2.active
    assert ws2.cell(1, 1).style_index == unlocked_idx
    assert ws2.cell(1, 2).style_index == hidden_idx


def test_granular_protection(tmp_path):
    """Test granular protection flags."""
    wb = Workbook()
    ws = wb.active

    ws.protect(insert_rows=True, delete_rows=False, select_locked_cells=False)

    prot = ws.protection
    assert prot["insert_rows"] is True
    assert prot["delete_rows"] is False
    assert prot["select_locked_cells"] is False
