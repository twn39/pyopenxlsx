import pytest
from pyopenxlsx import Workbook, load_workbook, load_workbook_async


def test_workbook_encryption(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1).value = "Top Secret Data"

    file_path = tmp_path / "encrypted.xlsx"

    # Save with password
    wb.save(str(file_path), password="my_secret_password")

    # Opening without password should fail or raise an error from C++
    with pytest.raises(Exception):
        load_workbook(str(file_path))

    # Opening with wrong password should fail
    with pytest.raises(Exception):
        load_workbook(str(file_path), password="wrong_password")

    # Opening with correct password should succeed
    wb2 = load_workbook(str(file_path), password="my_secret_password")
    assert wb2.active.cell(row=1, column=1).value == "Top Secret Data"


@pytest.mark.asyncio
async def test_workbook_encryption_async(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.cell(row=2, column=2).value = "Async Secret"

    file_path = tmp_path / "encrypted_async.xlsx"
    await wb.save_async(str(file_path), password="async_password")

    wb2 = await load_workbook_async(str(file_path), password="async_password")
    assert wb2.active.cell(row=2, column=2).value == "Async Secret"
