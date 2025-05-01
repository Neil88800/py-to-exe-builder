import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import threading
import time

class ExcelNameCounterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("姓名統計工具")
        self.root.geometry("550x360")
        self.root.configure(bg="#1e1e2f")

        self.filename = None
        self.save_path = tk.StringVar()
        self.output_filename = tk.StringVar(value="output.xlsx")

        self.set_theme()

        title = tk.Label(root, text="🔍 Excel 姓名統計工具", font=("Microsoft JhengHei UI", 16, "bold"),
                         fg="#00ffff", bg="#1e1e2f")
        title.pack(pady=15)

        self.select_btn = tk.Button(root, text="📂 選擇 Excel 檔", command=self.select_file, **self.button_style())
        self.select_btn.pack(pady=10)

        self.file_label = tk.Label(root, text="尚未選擇檔案", font=("Microsoft JhengHei UI", 10),
                                   fg="#cccccc", bg="#1e1e2f")
        self.file_label.pack()

        # 儲存路徑區塊
        path_frame = tk.Frame(root, bg="#1e1e2f")
        path_frame.pack(pady=12)
        tk.Label(path_frame, text="💾 儲存位置：", font=("Microsoft JhengHei UI", 10),
                 fg="#ffffff", bg="#1e1e2f").pack(side=tk.LEFT, padx=5)
        self.path_entry = tk.Entry(path_frame, textvariable=self.save_path, width=40, **self.entry_style())
        self.path_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(path_frame, text="選擇", command=self.select_save_path, **self.button_style(small=True)).pack(side=tk.LEFT)

        # 檔名設定區塊
        filename_frame = tk.Frame(root, bg="#1e1e2f")
        filename_frame.pack(pady=5)
        tk.Label(filename_frame, text="📝 檔名：", font=("Microsoft JhengHei UI", 10),
                 fg="#ffffff", bg="#1e1e2f").pack(side=tk.LEFT, padx=5)
        self.name_entry = tk.Entry(filename_frame, textvariable=self.output_filename, **self.entry_style())
        self.name_entry.pack(side=tk.LEFT)

        # 執行按鈕
        self.run_btn = tk.Button(root, text="⚙️ 開始處理", command=self.run_process, **self.button_style())
        self.run_btn.pack(pady=15)

        # 進度條
        self.progress = ttk.Progressbar(root, mode='indeterminate', length=400, style="Custom.Horizontal.TProgressbar")
        self.progress.pack(pady=5)

    def set_theme(self):
        # 自定義進度條樣式
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor='#2d2d3d',
                        background='#00ffff',
                        bordercolor='#1e1e2f',
                        lightcolor='#00ffff',
                        darkcolor='#00cccc',
                        thickness=10)

    def button_style(self, small=False):
        return {
            "font": ("Microsoft JhengHei UI", 9 if small else 10, "bold"),
            "fg": "#ffffff",
            "bg": "#005f73",
            "activebackground": "#008891",
            "activeforeground": "#ffffff",
            "relief": "flat",
            "bd": 0,
            "padx": 10,
            "pady": 5,
        }

    def entry_style(self):
        return {
            "font": ("Microsoft JhengHei UI", 10),
            "fg": "#ffffff",
            "bg": "#2b2b3d",
            "insertbackground": "#00ffff",
            "relief": "flat",
            "highlightthickness": 1,
            "highlightbackground": "#444444",
            "highlightcolor": "#00ffff",
        }

    def select_file(self):
        self.filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.filename:
            self.file_label.config(text=os.path.basename(self.filename))

    def select_save_path(self):
        directory = filedialog.askdirectory()
        if directory:
            self.save_path.set(directory)

    def run_process(self):
        if not self.filename:
            messagebox.showwarning("⚠️ 警告", "請先選擇一個 Excel 檔案。")
            return
        if not self.save_path.get():
            messagebox.showwarning("⚠️ 警告", "請選擇儲存位置。")
            return
        if not self.output_filename.get().strip():
            messagebox.showwarning("⚠️ 警告", "請輸入輸出檔名。")
            return

        threading.Thread(target=self.process_file).start()

    def process_file(self):
        self.progress.start()
        time.sleep(0.5)

        try:
            df = pd.read_excel(self.filename)
            name_counts = df.iloc[:, 0].value_counts().reset_index()
            name_counts.columns = ['姓名', '出現次數']

            output_path = os.path.join(
                self.save_path.get(),
                self.output_filename.get() if self.output_filename.get().endswith('.xlsx') else self.output_filename.get() + '.xlsx'
            )

            name_counts.to_excel(output_path, index=False)

            self.progress.stop()
            messagebox.showinfo("✅ 完成", f"處理完成！結果已儲存於：\n{output_path}")
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("❌ 錯誤", f"發生錯誤：\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelNameCounterApp(root)
    root.mainloop()
