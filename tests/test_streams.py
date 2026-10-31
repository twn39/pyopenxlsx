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


def test_stream_writer_phases_and_reader_sparse(tmp_path):
    import pytest

    file_path = tmp_path / "test_stream_advanced.xlsx"
    with Workbook() as wb:
        ws = wb.active
        with ws.stream_writer() as sw:
            # Phase 0: Column and pane configuration (fluent chaining)
            sw.set_column_width(1, 3, 20.0).set_column_visible(2, 2, True).freeze_panes(
                "B2"
            )

            # Phase 1: Rows
            sw.append_row(["Header 1", "Header 2", "Header 3"])
            sw.append_row([100, None, 300])

            # Phase 0 violation after rows started should fail-fast
            with pytest.raises(
                RuntimeError, match="must be configured before writing rows"
            ):
                sw.set_column_width(1, 1, 15.0)

            with pytest.raises(
                RuntimeError, match="must be configured before writing rows"
            ):
                sw.freeze_panes("A1")

            # Phase 2: Post-row decoration (fluent chaining)
            sw.merge_cells("A1:C1").auto_filter("A1:C2").add_hyperlink(
                "A2", "https://example.com", tooltip="Example"
            )

        wb.save(file_path)

    # Read back and verify sparse reading
    with Workbook(file_path) as wb:
        ws = wb.active
        with ws.stream_reader() as reader:
            # r1 has Header 1, Header 2, Header 3
            r1 = reader.next_row_sparse()
            assert len(r1) == 3
            assert r1[0]["value"] == "Header 1"

            # r2 has 100 at col 1 and 300 at col 3 (col 2 is None/empty)
            r2 = reader.next_row_sparse()
            # Sparse reading skips empty cell at col 2
            assert len(r2) == 2
            assert r2[0]["column"] == 1 and r2[0]["value"] == 100
            assert r2[1]["column"] == 3 and r2[1]["value"] == 300

    # Test iter_sparse convenience generator
    with Workbook(file_path) as wb:
        ws = wb.active
        with ws.stream_reader() as reader:
            sparse_rows = list(reader.iter_sparse())
            assert len(sparse_rows) == 2
            assert sparse_rows[1] == [(1, 100), (3, 300)]
