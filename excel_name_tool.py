# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
import time

class ExcelToolApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel 姓名統計工具")

        # 設定視窗大小與不可調整
        master.geometry("400x250")
        master.resizable(False, False)

        # 元件區
        self.label = tk.Label(master, text="請選擇含『工作表1』的 Excel 檔：")
        self.label.pack(pady=10)

        self.select_button = tk.Button(master, text="📂 選擇檔案", command=self.select_file)
        self.select_button.pack()

        self.file_label = tk.Label(master, text="", fg="blue")
        self.file_label.pack()

        self.save_button = tk.Button(master, text="💾 選擇儲存位置", command=self.select_save_path, state=tk.DISABLED)
        self.save_button.pack(pady=5)

        self.progress = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        self.start_button = tk.Button(master, text="🚀 開始執行", command=self.run_processing, state=tk.DISABLED)
        self.start_button.pack()

        # 路徑暫存
        self.file_path = None
        self.save_path = "產出報表.xlsx"

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="請選擇含有工作表1的 Excel 檔案",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"✔️ 已選檔：{file_path.split('/')[-1]}")
            self.save_button.config(state=tk.NORMAL)
        else:
            self.file_label.config(text="⚠️ 尚未選擇檔案")

    def select_save_path(self):
        save_path = filedialog.asksaveasfilename(
            title="儲存報表為...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if save_path:
            self.save_path = save_path
            self.start_button.config(state=tk.NORMAL)

    def run_processing(self):
        # 使用 Thread 避免 GUI 卡死
        threading.Thread(target=self.process_excel).start()

    def process_excel(self):
        try:
            self.progress["value"] = 10
            self.master.update_idletasks()
            time.sleep(0.2)

            xlsx = pd.ExcelFile(self.file_path)
            self.progress["value"] = 30
            self.master.update_idletasks()

            df1 = xlsx.parse('工作表1')
            self.progress["value"] = 50
            self.master.update_idletasks()

            data = df1.iloc[1:, 2].dropna().astype(str)
            names = data.str.extract(r'-(.+)$')[0].str.strip()
            summary = names.value_counts().reset_index()
            summary.columns = ['姓名', '次數']
            total = summary['次數'].sum()
            summary.loc[len(summary)] = ['總計', total]

            self.progress["value"] = 70
            self.master.update_idletasks()

            with pd.ExcelWriter(self.save_path, engine='openpyxl') as writer:
                summary.to_excel(writer, sheet_name='工作表2', startrow=2, startcol=1, index=False)
                ws = writer.sheets['工作表2']
                ws['B2'] = '姓名'
                ws['C2'] = '次數'

            self.progress["value"] = 100
            self.master.update_idletasks()
            time.sleep(0.2)
            messagebox.showinfo("完成", f"報表已成功儲存至：\n{self.save_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"處理失敗：\n{str(e)}")
        finally:
            self.progress["value"] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToolApp(root)
    root.mainloop()
