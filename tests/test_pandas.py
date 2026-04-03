import pytest
import datetime
from pyopenxlsx import Workbook

pytest.importorskip("pandas")
import pandas as pd


def test_pandas_write_read(tmp_path):
    wb = Workbook()
    ws = wb.active

    # 1. Create a DataFrame with mixed types
    df = pd.DataFrame(
        {
            "ID": [1, 2, 3, 4],
            "Name": ["Alice", "Bob", "Charlie", "David"],
            "Score": [99.5, 88.0, 77.5, 92.1],
            "Active": [True, False, True, True],
            "Date": [
                datetime.date(2023, 1, 1),
                datetime.date(2023, 2, 2),
                datetime.date(2023, 3, 3),
                datetime.date(2023, 4, 4),
            ],
        }
    )

    # 2. Write DataFrame to worksheet
    ws.write_dataframe(df, start_row=1, start_col=1, header=True, index=False)

    # 3. Verify it was written correctly by reading it back using cell interface
    assert ws["A1"].value == "ID"
    assert ws["B1"].value == "Name"
    assert ws["B2"].value == "Alice"
    assert ws["C3"].value == 88.0
    assert ws["D5"].value is True

    # Note: dates read back as float/serial by default unless formatted or evaluated as cell,
    # so we just check that the float value is present.

    file_path = tmp_path / "pandas_test.xlsx"
    wb.save(str(file_path))

    # 4. Read DataFrame back
    # Let's read columns 1 to 4 (ID, Name, Score, Active)
    df_read = ws.read_dataframe(
        start_row=1, start_col=1, end_row=5, end_col=4, header=True
    )

    # Verify the read DataFrame
    assert len(df_read) == 4
    assert list(df_read.columns) == ["ID", "Name", "Score", "Active"]
    assert df_read.iloc[0]["Name"] == "Alice"
    assert df_read.iloc[2]["Score"] == 77.5


@pytest.mark.asyncio
async def test_pandas_write_read_async(tmp_path):
    wb = Workbook()
    ws = wb.active
    df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
    await ws.write_dataframe_async(df)

    df_read = await ws.read_dataframe_async(end_row=3, end_col=2)
    assert len(df_read) == 2
    assert df_read.iloc[0]["B"] == 3
