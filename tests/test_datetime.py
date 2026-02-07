from datetime import datetime, date
from pyopenxlsx import Workbook


def test_datetime_write_read(tmp_path):
    wb = Workbook()
    ws = wb.active

    # Test writing datetime
    dt = datetime(2023, 10, 27, 14, 30, 0)
    ws["A1"].value = dt

    # Verify value is float initially if style not set?
    # Actually, XLCell setter sets float serial.
    # Cell.value getter converts IF style is date.
    # But writing value doesn't automatically set default date style in current implementation?
    # If not, Cell.value will return float.
    # Let's check if we need to manually set style or if we expect auto-style (not implemented yet).
    # For now, let's manually set style to verify conversion.

    wb.add_style(number_format="yyyy-mm-dd hh:mm:ss")  # logic might not auto-apply?
    # wb.add_style returns index.
    style_idx = wb.add_style(number_format="yyyy-mm-dd hh:mm:ss")
    ws["A1"].style_index = style_idx

    # Now read back
    # Need to save and reload to ensure full cycle or just read from memory?
    # pyopenxlsx works on in-memory object too.
    read_dt = ws["A1"].value
    assert isinstance(read_dt, datetime)
    assert read_dt.year == 2023
    # openxlsx might have precision issues or timezone?
    # assertions should be close enough or exact seconds.
    assert abs((read_dt - dt).total_seconds()) < 1.0

    output = tmp_path / "test_datetime.xlsx"
    wb.save(str(output))
    assert output.exists()


def test_date_write_read(tmp_path):
    wb = Workbook()
    ws = wb.active

    d = date(2023, 12, 25)
    ws["B2"].value = d

    # Set date style
    style_idx = wb.add_style(number_format="yyyy-mm-dd")
    ws["B2"].style_index = style_idx

    read_d = ws["B2"].value
    # Note: XLDateTime usually returns datetime.
    # Our wrapper converts using as_datetime() which returns datetime.
    # So we expect datetime, even if we wrote date.
    assert isinstance(read_d, datetime)
    assert read_d.year == 2023
    assert read_d.month == 12
    assert read_d.day == 25
