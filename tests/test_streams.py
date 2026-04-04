from pyopenxlsx import Workbook, XLFont


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
        
        # If the reader starts at row 1 but we haven't read anything yet
        # current_row might return 0 or 1 depending on implementation
        # Let's check what it actually does
        print(f"Initial index: {reader.current_row_index}")
        
        assert next(reader) == [1]
        assert reader.current_row_index == 1
        assert next(reader) == [2]
        assert reader.current_row_index == 2
        assert next(reader) == [3]
        assert reader.current_row_index == 3
        assert not reader.has_next()
