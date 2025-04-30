# -*- coding: utf-8 -*-
import pandas as pd

def process_excel(input_file, output_file):
    """
    這個函數將處理 Excel 檔案，擷取並統計姓名，並將結果保存為新的 Excel 檔案。
    :param input_file: 原始 Excel 檔案路徑，必須包含名為「工作表1」的工作表
    :param output_file: 輸出的 Excel 檔案路徑
    """
    # 讀取工作表1
    xlsx = pd.ExcelFile(input_file)
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

    # 產出新的 Excel 並儲存
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='工作表2', startrow=2, startcol=1, index=False)
        ws = writer.sheets['工作表2']
        ws['B2'] = '姓名'
        ws['C2'] = '次數'

    print(f"檔案已儲存為: {output_file}")

if __name__ == "__main__":
    input_file = input("請輸入原始 Excel 檔案路徑 (包含工作表1): ")
    output_file = input("請輸入輸出檔案路徑 (例如: 產出報表.xlsx): ")
    process_excel(input_file, output_file)
