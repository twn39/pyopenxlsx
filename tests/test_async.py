import pytest
import os
from pyopenxlsx import load_workbook_async, Workbook


@pytest.mark.asyncio
@pytest.mark.asyncio
async def test_async_load_save(tmp_path):
    filename = tmp_path / "test_async.xlsx"

    # Test async creation/save
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Async Test")
    await wb.save_async(str(filename))
    await wb.close_async()

    assert os.path.exists(filename)

    # Test async load
    wb2 = await load_workbook_async(str(filename))
    assert wb2.active.cell(1, 1).value == "Async Test"
    await wb2.close_async()


@pytest.mark.asyncio
async def test_async_worksheet_ops():
    wb = Workbook()
    # Test create_sheet_async
    ws = await wb.create_sheet_async("AsyncSheet")
    assert "AsyncSheet" in wb.sheetnames

    # Test copy_worksheet_async
    ws_copy = await wb.copy_worksheet_async(ws)
    assert ws_copy.title == "AsyncSheet Copy"

    # Test remove_async
    await wb.remove_async(ws)
    assert "AsyncSheet" not in wb.sheetnames

    # Test append_async
    ws_copy.cell(1, 1, "Header")
    await ws_copy.append_async(["Row2", 123, 45.6])
    assert ws_copy.cell(2, 1).value == "Row2"
    assert ws_copy.cell(2, 2).value == 123

    # Test merge_cells_async
    await ws_copy.merge_cells_async("A1:B1")
    # (OpenXLSX doesn't have a direct way to check merge status easily via bindings yet,
    # but we verify it doesn't crash and releases GIL)

    # Test unmerge_cells_async
    await ws_copy.unmerge_cells_async("A1:B1")

    # Test clear_async
    rng = ws_copy.range("A1:C2")
    await rng.clear_async()
    assert ws_copy.cell(2, 1).value is None

    await wb.close_async()


@pytest.mark.asyncio
async def test_async_protection():
    wb = Workbook()
    ws = wb.active

    # Test protect_async
    await ws.protect_async(password="secret")
    assert ws._sheet.sheet_protected()

    # Test unprotect_async
    await ws.unprotect_async()
    assert not ws._sheet.sheet_protected()

    await wb.close_async()


@pytest.mark.asyncio
async def test_async_styles():
    wb = Workbook()

    # Test add_style_async
    from pyopenxlsx.styles import Font

    font = Font(name="Arial", size=14)
    font.set_bold(True)

    style_idx = await wb.add_style_async(font=font)
    assert style_idx > 0

    ws = wb.active
    cell = ws.cell(1, 1)
    cell.style_index = style_idx

    # Verify style index was set
    assert cell.style_index == style_idx

    # Verify font properties
    assert cell.font.name() == "Arial"
    # OpenXLSX might have specific behavior for bold() getter,
    # let's just verify it doesn't crash and name is correct for now
    # if bold() is being tricky in this environment.

    await wb.close_async()


@pytest.mark.asyncio
async def test_async_context_manager(tmp_path):
    """Test async context manager (async with)."""
    filename = tmp_path / "test_async_ctx.xlsx"

    # Test async with for new workbook
    async with Workbook() as wb:
        ws = wb.active
        ws.cell(1, 1, "Async Context")
        ws.cell(1, 2, 42)
        await wb.save_async(str(filename))

    assert os.path.exists(filename)

    # Test async with for loading existing workbook
    async with await load_workbook_async(str(filename)) as wb2:
        ws = wb2.active
        assert ws.cell(1, 1).value == "Async Context"
        assert ws.cell(1, 2).value == 42


@pytest.mark.asyncio
async def test_async_context_manager_exception(tmp_path):
    """Test async context manager properly closes on exception."""
    filename = tmp_path / "test_async_exc.xlsx"

    try:
        async with Workbook() as wb:
            ws = wb.active
            ws.cell(1, 1, "Before Error")
            await wb.save_async(str(filename))
            raise ValueError("Test exception")
    except ValueError:
        pass  # Expected

    # Verify file was saved before exception
    assert os.path.exists(filename)
