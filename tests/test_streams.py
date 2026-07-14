from pyopenxlsx import Workbook
from pyopenxlsx._openxlsx import (
    XLFont,
    XLStreamEmptyRowPolicy,
    XLStreamReadOptions,
    XLStreamRowOpts,
)


def test_worksheet_streams(tmp_path):
    file_path = tmp_path / "test_streams_v2.xlsx"

    with Workbook() as wb:
        ws = wb.active

        # Add a style correctly
        f = XLFont()
        f.set_bold(True)
        style_idx = wb.add_style(font=wb.styles.fonts().create(f))

        # Use context manager for stream_writer
        with ws.stream_writer() as writer:
            assert writer.is_active
            writer.append_row([1, "Test", 3.14])
            writer.append_row([(2, style_idx), "Data", 2.71])

        # Stream should be closed now
        assert not writer.is_active

        wb.save(file_path)

    # Reading back using iterator
    with Workbook(file_path) as wb:
        ws = wb.active
        reader = ws.stream_reader()

        rows = list(reader)

        assert len(rows) == 2
        assert rows[0] == [1, "Test", 3.14]
        assert rows[1] == [2, "Data", 2.71]


def test_stream_reader_index(tmp_path):
    file_path = tmp_path / "test_index.xlsx"
    with Workbook() as wb:
        ws = wb.active
        with ws.stream_writer() as writer:
            writer.append_row([1])
            writer.append_row([2])
            writer.append_row([3])
        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        reader = ws.stream_reader()

        assert next(reader) == [1]
        assert reader.current_row_index == 1
        assert next(reader) == [2]
        assert reader.current_row_index == 2
        assert next(reader) == [3]
        assert reader.current_row_index == 3
        assert not reader.has_next()


def test_stream_set_row_and_opts(tmp_path):
    file_path = tmp_path / "test_set_row.xlsx"
    with Workbook() as wb:
        ws = wb.active
        opts = XLStreamRowOpts()
        opts.height = 30.0
        opts.hidden = False
        with ws.stream_writer() as writer:
            writer.set_row(1, 1, ["A", "B"])
            writer.set_row_ref("A3", [10, 20], opts)
            assert writer.last_row == 3
            assert writer.max_column >= 2
        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        opts = XLStreamReadOptions()
        opts.empty_rows = XLStreamEmptyRowPolicy.EmitEmptyRows
        with ws.stream_reader(options=opts) as reader:
            r1 = reader.next_row()
            assert r1[:2] == ["A", "B"]
            assert reader.current_row_index == 1
            # synthetic empty row 2
            reader.next_row()
            assert reader.current_row_index == 2
            assert reader.current_row_opts().get("is_synthetic_empty") is True
            r3 = reader.next_row()
            assert r3[:2] == [10, 20]
            assert reader.current_row_index == 3


def test_stream_formula_and_detailed(tmp_path):
    file_path = tmp_path / "test_stream_formula.xlsx"
    with Workbook() as wb:
        ws = wb.active
        with ws.stream_writer() as writer:
            writer.append_row([1, 2, {"value": 3, "formula": "A1+B1"}])
        wb.save(file_path)

    with Workbook(file_path) as wb:
        ws = wb.active
        with ws.stream_reader() as reader:
            detailed = reader.next_row_detailed()
            assert len(detailed) >= 3
            # formula cell should expose formula text when present
            formulas = [c.get("formula") for c in detailed if c.get("formula")]
            assert any(f and "A1" in f for f in formulas)
