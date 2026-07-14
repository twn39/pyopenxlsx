from datetime import datetime, date
from pyopenxlsx import Workbook


def test_datetime_write_read(tmp_path):
    wb = Workbook()
    ws = wb.active

    # Test writing datetime — auto_date_formats (default True) applies number format
    dt = datetime(2023, 10, 27, 14, 30, 0)
    ws["A1"].value = dt

    read_dt = ws["A1"].value
    assert isinstance(read_dt, datetime)
    assert read_dt.year == 2023
    assert abs((read_dt - dt).total_seconds()) < 1.0

    output = tmp_path / "test_datetime.xlsx"
    wb.save(str(output))
    assert output.exists()
    wb.close()


def test_date_write_read(tmp_path):
    wb = Workbook()
    ws = wb.active

    d = date(2023, 12, 25)
    ws["B2"].value = d

    read_d = ws["B2"].value
    # Wrapper converts serial + date format back to datetime
    assert isinstance(read_d, datetime)
    assert read_d.year == 2023
    assert read_d.month == 12
    assert read_d.day == 25
    wb.close()


def test_auto_date_formats_opt_out():
    wb = Workbook()
    wb.auto_date_formats = False
    ws = wb.active
    dt = datetime(2023, 10, 27, 14, 30, 0)
    ws["A1"].value = dt
    # Without auto formats, value remains a float serial
    assert isinstance(ws["A1"].value, float)
    wb.close()
