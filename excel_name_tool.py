import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import threading
import time

class ExcelNameCounterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("統計工具 Pro")
        self.root.geometry("640x480")
        self.root.configure(bg="#0f0f1a")
        self.root.resizable(False, False)

        self.filename = None
        self.save_path = tk.StringVar()
        self.output_filename = tk.StringVar()

        self.set_theme()

        title = tk.Label(root, text="報表用統計工具", font=("Segoe UI", 20, "bold"),
                         fg="#00f0ff", bg="#0f0f1a")
        title.pack(pady=20)

        self.select_btn = tk.Button(root, text="📂 載入 Excel 檔案", command=self.select_file, **self.button_style(text_color="#000000"))
        self.select_btn.pack(pady=10)

        self.file_label = tk.Label(root, text="尚未選擇檔案", font=("Segoe UI", 12),
                                   fg="#aaaaaa", bg="#0f0f1a")
        self.file_label.pack()

        path_frame = tk.Frame(root, bg="#0f0f1a")
        path_frame.pack(pady=15)
        tk.Label(path_frame, text="📁 儲存路徑：", font=("Segoe UI", 12),
                 fg="#ffffff", bg="#0f0f1a").pack(side=tk.LEFT, padx=5)
        self.path_entry = tk.Entry(path_frame, textvariable=self.save_path, width=40, **self.entry_style())
        self.path_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(path_frame, text="選擇", command=self.select_save_path, **self.button_style(small=True, text_color="#000000")).pack(side=tk.LEFT)

        filename_frame = tk.Frame(root, bg="#0f0f1a")
        filename_frame.pack(pady=5)
        tk.Label(filename_frame, text="📝 檔名（不需輸入副檔名）：", font=("Segoe UI", 12),
                 fg="#ffffff", bg="#0f0f1a").pack(side=tk.LEFT, padx=5)
        self.name_entry = tk.Entry(filename_frame, textvariable=self.output_filename, **self.entry_style())
        self.name_entry.pack(side=tk.LEFT)

        self.run_btn = tk.Button(root, text="🚀 開始分析", command=self.run_process, **self.button_style(text_color="#000000"))
        self.run_btn.pack(pady=20)

        self.progress = ttk.Progressbar(root, mode='determinate', length=440, style="Custom.Horizontal.TProgressbar")
        self.progress.pack(pady=8)

        self.progress_info = tk.Label(root, text="", font=("Segoe UI", 10), fg="#00f0ff", bg="#0f0f1a")
        self.progress_info.pack()

        self.footer = tk.Label(root, text="🔧 Powered by Pandas & Tkinter | Abby 專用版", font=("Segoe UI", 10),
                               fg="#5555aa", bg="#0f0f1a")
        self.footer.pack(pady=12)

    def set_theme(self):
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor='#1a1a2e',
                        background='#00f0ff',
                        bordercolor='#0f0f1a',
                        lightcolor='#00e0dd',
                        darkcolor='#00e0dd',
                        thickness=12)

    def button_style(self, small=False, text_color="#ffffff"):
        return {
            "font": ("Segoe UI", 10 if small else 14, "bold"),
            "fg": text_color,
            "bg": "#0077b6",
            "activebackground": "#00b4d8",
            "activeforeground": text_color,
            "relief": "flat",
            "bd": 0,
            "padx": 14,
            "pady": 7,
        }

    def entry_style(self):
        return {
            "font": ("Segoe UI", 12),
            "fg": "#ffffff",
            "bg": "#2a2a40",
            "insertbackground": "#00f0ff",
            "relief": "flat",
            "highlightthickness": 1,
            "highlightbackground": "#444444",
            "highlightcolor": "#00f0ff",
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
            messagebox.showwarning("⚠️ 錯誤", "請先選擇一個 Excel 檔案。")
            return
        if not self.save_path.get():
            messagebox.showwarning("⚠️ 錯誤", "請選擇儲存位置。")
            return
        if not self.output_filename.get().strip():
            messagebox.showwarning("⚠️ 錯誤", "請輸入輸出檔名。")
            return

        threading.Thread(target=self.process_file).start()

    def process_file(self):
        try:
            df = pd.read_excel(self.filename)
            total = len(df)

            name_counts = df.iloc[:, 0].value_counts().reset_index()
            name_counts.columns = ['姓名', '出現次數']

            filename = self.output_filename.get().strip()
            if not filename.endswith(".xlsx"):
                filename += ".xlsx"

            output_path = os.path.join(self.save_path.get(), filename)

            # 模擬分段進度
            self.progress["value"] = 0
            for i in range(1, 101):
                time.sleep(0.01)  # 模擬運算延遲
                self.progress["value"] = i
                remaining = (100 - i) * 0.01
                self.progress_info.config(text=f"處理進度：{i}% ｜ 預估剩餘 {remaining:.1f} 秒")
                self.root.update_idletasks()

            name_counts.to_excel(output_path, index=False)

            self.progress_info.config(text="✅ 分析完成！")
            messagebox.showinfo("✅ 完成", f"處理完成！結果已儲存於：\n{output_path}")
        except Exception as e:
            self.progress_info.config(text="❌ 分析失敗")
            messagebox.showerror("❌ 錯誤", f"發生錯誤：\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelNameCounterApp(root)
    root.mainloop()
