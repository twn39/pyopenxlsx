from pyopenxlsx import Workbook


def test_worksheet_streams(tmp_path):
    file_path = tmp_path / "test_streams.xlsx"

    with Workbook() as wb:
        ws = wb.active
        writer = ws.stream_writer()
        assert writer.is_stream_active()

        # Add a style
        style_idx = wb.add_style(font=wb.styles.fonts().create())  # dummy style

        writer.append_row([1, "Test", 3.14])
        writer.append_row([(2, style_idx), "Data", 2.71])
        writer.close()

        wb.save(file_path)

    print("Reading")
    with Workbook(file_path) as wb:
        ws = wb.active
        reader = ws.stream_reader()

        rows = []
        while reader.has_next():
            print("Row:", reader.current_row())
            row = reader.next_row()
            rows.append(row)

        print("Done")
        assert len(rows) == 2
        assert rows[1][0] == 2
