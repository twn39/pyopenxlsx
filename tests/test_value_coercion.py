"""Matrix tests: Cell vs bulk write paths share value/date coercion."""

from datetime import date, datetime

import pytest

from pyopenxlsx import Workbook
from pyopenxlsx._coercion import (
    DATE_KIND_DATE,
    DATE_KIND_DATETIME,
    coerce_cell_value,
    datetime_to_serial,
)


@pytest.fixture
def wb_ws():
    wb = Workbook()
    ws = wb.active
    yield wb, ws
    wb.close()


def _assert_date_like(cell_value, expected_date, *, has_time: bool):
    assert isinstance(cell_value, datetime)
    assert cell_value.year == expected_date.year
    assert cell_value.month == expected_date.month
    assert cell_value.day == expected_date.day
    if has_time:
        assert cell_value.hour == expected_date.hour
        assert cell_value.minute == expected_date.minute


class TestCoerceHelpers:
    def test_coerce_date_and_datetime(self):
        d = date(2023, 12, 25)
        dt = datetime(2023, 10, 27, 14, 30, 0)
        serial_d, kind_d = coerce_cell_value(d)
        serial_dt, kind_dt = coerce_cell_value(dt)
        assert kind_d == DATE_KIND_DATE
        assert kind_dt == DATE_KIND_DATETIME
        assert serial_d == pytest.approx(datetime_to_serial(d))
        assert serial_dt == pytest.approx(datetime_to_serial(dt))

    def test_coerce_passthrough(self):
        for val in (None, 1, 1.5, True, "hello"):
            out, kind = coerce_cell_value(val)
            assert out is val or out == val
            assert kind is None


class TestPathMatrix:
    """Same values written via every public write path should round-trip alike."""

    @pytest.mark.parametrize(
        "writer",
        [
            "cell",
            "set_cell_value",
            "write_row",
            "write_rows",
            "set_cells",
        ],
    )
    def test_datetime_auto_format(self, wb_ws, writer):
        wb, ws = wb_ws
        dt = datetime(2023, 10, 27, 14, 30, 0)

        if writer == "cell":
            ws.cell(1, 1).value = dt
        elif writer == "set_cell_value":
            ws.set_cell_value(1, 1, dt)
        elif writer == "write_row":
            ws.write_row(1, [dt])
        elif writer == "write_rows":
            ws.write_rows(1, [[dt]])
        elif writer == "set_cells":
            ws.set_cells([(1, 1, dt)])

        cell = ws.cell(1, 1)
        assert cell.is_date is True
        _assert_date_like(cell.value, dt, has_time=True)

    @pytest.mark.parametrize(
        "writer",
        [
            "cell",
            "set_cell_value",
            "write_row",
            "write_rows",
            "set_cells",
            "append_row",
        ],
    )
    def test_date_auto_format(self, wb_ws, writer):
        wb, ws = wb_ws
        d = date(2023, 12, 25)

        if writer == "cell":
            ws.cell(2, 2).value = d
        elif writer == "set_cell_value":
            ws.set_cell_value(2, 2, d)
        elif writer == "write_row":
            ws.write_row(2, [None, d])
        elif writer == "write_rows":
            ws.write_rows(2, [[None, d]])
        elif writer == "set_cells":
            ws.set_cells([(2, 2, d)])
        elif writer == "append_row":
            # Empty sheet: first append is row 1; use that and check col 1
            ws.append_row([d])
            cell = ws.cell(1, 1)
            assert cell.is_date is True
            _assert_date_like(cell.value, d, has_time=False)
            return

        cell = ws.cell(2, 2)
        assert cell.is_date is True
        _assert_date_like(cell.value, d, has_time=False)

    @pytest.mark.parametrize(
        "writer",
        ["cell", "set_cell_value", "write_rows", "set_cells"],
    )
    def test_auto_date_formats_opt_out(self, writer):
        wb = Workbook()
        wb.auto_date_formats = False
        ws = wb.active
        dt = datetime(2023, 10, 27, 14, 30, 0)

        if writer == "cell":
            ws.cell(1, 1).value = dt
        elif writer == "set_cell_value":
            ws.set_cell_value(1, 1, dt)
        elif writer == "write_rows":
            ws.write_rows(1, [[dt]])
        elif writer == "set_cells":
            ws.set_cells([(1, 1, dt)])

        # Without auto formats, raw serial float (not datetime wrapper)
        assert isinstance(ws.cell(1, 1).value, float)
        assert ws.cell(1, 1).is_date is False
        wb.close()

    def test_mixed_row_types(self, wb_ws):
        wb, ws = wb_ws
        d = date(2024, 1, 15)
        dt = datetime(2024, 1, 15, 9, 0, 0)
        ws.write_rows(
            1,
            [
                ["name", "when", "flag", "n"],
                ["a", d, True, 1],
                ["b", dt, False, 2.5],
                ["c", None, None, None],
            ],
        )
        assert ws.cell(1, 1).value == "name"
        assert ws.cell(2, 1).value == "a"
        assert ws.cell(2, 2).is_date is True
        _assert_date_like(ws.cell(2, 2).value, d, has_time=False)
        assert ws.cell(3, 2).is_date is True
        _assert_date_like(ws.cell(3, 2).value, dt, has_time=True)
        assert ws.cell(2, 3).value is True
        assert ws.cell(3, 4).value == pytest.approx(2.5)
        assert ws.cell(4, 2).value is None

    def test_does_not_overwrite_existing_date_style(self, wb_ws):
        wb, ws = wb_ws
        custom = wb.add_style(number_format="dd/mm/yyyy")
        ws.cell(1, 1).style_index = custom
        d = date(2020, 6, 1)
        ws.set_cell_value(1, 1, d)
        # Style should remain the custom one (already a date format)
        assert ws.cell(1, 1).style_index == custom
        assert ws.cell(1, 1).is_date is True
