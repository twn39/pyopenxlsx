import datetime
import pandas as pd
import numpy as np
from pyopenxlsx import Workbook, Font, Fill

def create_pandas_demo():
    print("🚀 正在生成 DataFrame Excel...")
    
    # 1. 准备一份混合了各种数据类型的大型 DataFrame
    # 包含：整数、字符串、浮点数、布尔值、日期、空值(NaN)
    data = {
        "User ID": range(1001, 10001),  # 9000 行整数
        "Department": ["Sales", "Engineering", "HR"] * 3000, # 字符串
        "Salary (USD)": np.random.uniform(45000, 150000, 9000).round(2), # 浮点数
        "Is Manager": [True, False, False] * 3000, # 布尔值
        "Hire Date": [datetime.date(2020, 1, 15) + datetime.timedelta(days=i) for i in range(9000)], # 日期
        "Bonus": [np.nan if i % 10 == 0 else 5000.0 for i in range(9000)] # 包含空值的浮点数
    }
    
    df = pd.DataFrame(data)
    
    # 2. 写入 Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Pandas_Export"
    
    # 手动加点样式让它好看点
    # 给 Salary 和 Bonus 列加上货币格式
    currency_style = wb.add_style(number_format="$#,##0.00")
    
    # 给 Date 列加上日期格式
    date_style = wb.add_style(number_format="yyyy-mm-dd")
    
    # 一键将 DataFrame 写入 A1 开始的位置，并且利用 stream_writer 的底层逻辑同时极速赋予样式！
    ws.write_dataframe(df, start_row=1, start_col=1, header=True, index=False, column_styles={
        "Salary (USD)": currency_style,
        "Hire Date": date_style,
        "Bonus": currency_style
    })
    
    # 给表头加上蓝色背景和白色粗体
    header_style = wb.add_style(
        font=Font(bold=True, color="FFFFFF"),
        fill=Fill(pattern_type="solid", color="4F81BD")
    )
    for col in range(1, 7):
        ws.cell(row=1, column=col).style_index = header_style
        ws.auto_fit_column(col) # 自动适应列宽
        
    # 3. 保存文件
    file_name = "pandas_dataframe_demo.xlsx"
    wb.save(file_name)
    print(f"✅ 成功生成 '{file_name}' (9000行 x 6列)！")
    
    # 4. 验证读取功能
    print("\n🔍 正在测试从生成的 Excel 读取回 DataFrame...")
    # 只读取前 5 行数据 (包含表头，所以 end_row=6)
    df_read = ws.read_dataframe(start_row=1, start_col=1, end_row=6, end_col=6, header=True)
    
    # 手动将 Excel 序列号浮点数转回日期 (这是高性能读取的常见操作)
    df_read["Hire Date"] = pd.to_datetime(df_read["Hire Date"], unit='D', origin='1899-12-30').dt.date
    
    print("\n读取到的前 5 行数据：")
    print(df_read)
    print("\n✅ 读取测试通过！")

if __name__ == "__main__":
    create_pandas_demo()
