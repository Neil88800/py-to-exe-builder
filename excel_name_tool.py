# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import io

# 讓 tkinter 啟動並隱藏主視窗
root = tk.Tk()
root.withdraw()  # 隱藏主視窗

# 使用 filedialog 讓使用者選擇 Excel 檔案
file_path = filedialog.askopenfilename(
    title="請選擇含有工作表1的 Excel 檔案",
    filetypes=[("Excel Files", "*.xlsx;*.xls")]
)

if not file_path:
    print("未選擇檔案，程序結束。")
else:
    # 開啟選擇的 Excel 檔案
    xlsx = pd.ExcelFile(file_path)

    # 讀取工作表1
    df1 = xlsx.parse('工作表1')

    # 擷取 C 欄從第 2 列開始的資料
    data = df1.iloc[1:, 2].dropna().astype(str)
    names = data.str.extract(r'-(.+)$')[0].str.strip()

    # 統計並排序
    summary = names.value_counts().reset_index()
    summary.columns = ['姓名', '次數']

    # 加總
    total = summary['次數'].sum()
    summary.loc[len(summary)] = ['總計', total]

    # 產出新的 Excel 並下載
    output_file = "產出報表.xlsx"
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='工作表2', startrow=2, startcol=1, index=False)
        ws = writer.sheets['工作表2']
        ws['B2'] = '姓名'
        ws['C2'] = '次數'

    # 提示使用者檔案已儲存
    print(f"檔案已儲存為 {output_file}")
