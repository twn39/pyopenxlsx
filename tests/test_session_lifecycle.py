import gc
import os
import pytest
import numpy as np
from pyopenxlsx import Workbook
from pyopenxlsx._session import _DocumentSession


def test_session_temp_file_cleanup_on_close():
    wb = Workbook()
    temp_path = wb._temp_file
    assert temp_path is not None
    assert os.path.exists(temp_path)

    wb.close()
    assert not os.path.exists(temp_path)
    assert wb._closed


def test_session_temp_file_cleanup_on_gc():
    session = _DocumentSession()
    temp_path = session.temp_file
    assert temp_path is not None
    assert os.path.exists(temp_path)

    # Delete session and force garbage collection
    del session
    gc.collect()

    assert not os.path.exists(temp_path)


def test_closed_session_guards_all_operations():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "test")

    wb.close()
    assert wb._closed
    assert ws._closed

    with pytest.raises(ValueError, match="closed"):
        wb.save("test.xlsx")

    with pytest.raises(ValueError, match="closed"):
        wb.compact_styles()

    with pytest.raises(ValueError, match="closed"):
        ws.cell(1, 1)

    with pytest.raises(ValueError, match="closed"):
        _ = ws["A1"]

    with pytest.raises(ValueError, match="closed"):
        ws.set_cell_value(1, 1, 42)

    with pytest.raises(ValueError, match="closed"):
        ws.get_cell_value(1, 1)

    with pytest.raises(ValueError, match="closed"):
        ws.get_range_data(1, 1, 2, 2)

    with pytest.raises(ValueError, match="closed"):
        ws.get_rows_data()

    with pytest.raises(ValueError, match="closed"):
        ws.get_row_values(1)

    with pytest.raises(ValueError, match="closed"):
        list(ws.iter_row_values())

    with pytest.raises(ValueError, match="closed"):
        ws.write_range(1, 1, np.zeros((2, 2)))

    with pytest.raises(ValueError, match="closed"):
        ws.get_range_values(1, 1, 2, 2)

    with pytest.raises(ValueError, match="closed"):
        ws.range("A1:B2")

    with pytest.raises(ValueError, match="closed"):
        ws.merge_cells("A1:B2")

    with pytest.raises(ValueError, match="closed"):
        ws.stream_writer()

    with pytest.raises(ValueError, match="closed"):
        ws.stream_reader()


def test_worksheet_keeps_workbook_session_alive():
    # If user extracts worksheet directly without storing wb
    ws = Workbook().active
    ws.cell(1, 1, 100)
    assert ws.get_cell_value(1, 1) == 100
    assert not ws._closed
    # ws._workbook retains reference to wb and its session
    assert ws._workbook is not None
    ws._workbook.close()
    assert ws._closed
    with pytest.raises(ValueError, match="closed"):
        ws.cell(1, 1)
