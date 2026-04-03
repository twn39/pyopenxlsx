import datetime
import time
import pandas as pd
import numpy as np
from pyopenxlsx import Workbook, Font, Fill

def test_1m_rows():
    print("🚀 Generating DataFrame with 1,000,000 rows...")
    t0 = time.time()
    
    rows = 1_000_000
    df = pd.DataFrame({
        "ID": np.arange(1, rows + 1),
        "Department": np.random.choice(["Sales", "Engineering", "HR"], rows),
        "Salary": np.random.uniform(50000, 150000, rows).round(2),
        "Is Active": np.random.choice([True, False], rows),
        "Join Date": [datetime.date(2020, 1, 1)] * rows  # Simplified for fast generation
    })
    
    t1 = time.time()
    print(f"✅ DataFrame generated in {t1 - t0:.2f}s")
    
    print("🚀 Writing 1,000,000 rows to pyopenxlsx...")
    wb = Workbook()
    ws = wb.active
    ws.title = "1M_Rows"
    
    currency_style = wb.add_style(number_format="$#,##0.00")
    date_style = wb.add_style(number_format="yyyy-mm-dd")
    
    t2 = time.time()
    ws.write_dataframe(df, header=True, index=False, column_styles={
        "Salary": currency_style,
        "Join Date": date_style
    })
    t3 = time.time()
    print(f"✅ Wrote to C++ DOM/Stream in {t3 - t2:.2f}s")
    
    print("🚀 Saving to disk (XML compression)...")
    file_name = "1m_rows_demo.xlsx"
    wb.save(file_name)
    t4 = time.time()
    print(f"✅ Saved '{file_name}' in {t4 - t3:.2f}s")
    print(f"🎉 Total Excel Time (Write + Save): {t4 - t2:.2f}s")

if __name__ == "__main__":
    test_1m_rows()
