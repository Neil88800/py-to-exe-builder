# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
import time
import os
from datetime import datetime
from tkinter.font import Font
import webbrowser

class ModernExcelToolApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel 姓名統計工具")
        
        # 設定視窗大小與圖示
        master.geometry("500x400")
        master.resizable(False, False)
        
        # 嘗試設定圖示 (如果有的話)
        try:
            master.iconbitmap("app_icon.ico")  # 你需要準備一個圖示檔案
        except:
            pass
            
        # 設定顏色主題
        self.bg_color = "#f5f7fa"
        self.accent_color = "#4a6ee0"
        self.button_color = "#4a6ee0"
        self.button_text_color = "white"
        self.text_color = "#333333"
        
        # 設定主視窗背景
        master.configure(bg=self.bg_color)
        
        # 建立自訂字型
        self.title_font = Font(family="Microsoft JhengHei UI", size=14, weight="bold")
        self.normal_font = Font(family="Microsoft JhengHei UI", size=10)
        self.button_font = Font(family="Microsoft JhengHei UI", size=10, weight="bold")
        
        # 建立主框架
        self.main_frame = tk.Frame(master, bg=self.bg_color, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 標題區
        self.title_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        self.title_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.title_label = tk.Label(
            self.title_frame, 
            text="Excel 姓名統計工具", 
            font=self.title_font, 
            bg=self.bg_color, 
            fg=self.accent_color
        )
        self.title_label.pack()
        
        self.subtitle_label = tk.Label(
            self.title_frame, 
            text="從 Excel 檔案中提取並統計姓名出現次數", 
            font=self.normal_font, 
            bg=self.bg_color, 
            fg=self.text_color
        )
        self.subtitle_label.pack(pady=(5, 0))
        
        # 檔案選擇區
        self.file_frame = tk.LabelFrame(
            self.main_frame, 
            text="檔案選擇", 
            font=self.normal_font, 
            bg=self.bg_color, 
            fg=self.text_color,
            padx=15, 
            pady=15
        )
        self.file_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.file_label = tk.Label(
            self.file_frame, 
            text="請選擇含有「工作表1」的 Excel 檔案：", 
            font=self.normal_font, 
            bg=self.bg_color, 
            fg=self.text_color
        )
        self.file_label.pack(anchor=tk.W)
        
        self.file_path_var = tk.StringVar()
        self.file_path_entry = tk.Entry(
            self.file_frame, 
            textvariable=self.file_path_var, 
            font=self.normal_font,
            width=40,
            state="readonly"
        )
        self.file_path_entry.pack(fill=tk.X, pady=(5, 0))
        
        self.select_button = tk.Button(
            self.file_frame, 
            text="📂 選擇檔案", 
            command=self.select_file,
            font=self.button_font,
            bg=self.button_color,
            fg=self.button_text_color,
            activebackground=self.accent_color,
            activeforeground=self.button_text_color,
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        self.select_button.pack(pady=(10, 0))
        
        # 儲存選項區
        self.save_frame = tk.LabelFrame(
            self.main_frame, 
            text="儲存選項", 
            font=self.normal_font, 
            bg=self.bg_color, 
            fg=self.text_color,
            padx=15, 
            pady=15
        )
        self.save_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.save_path_var = tk.StringVar()
        self.save_path_var.set("產出報表.xlsx")
        self.save_path_entry = tk.Entry(
            self.save_frame, 
            textvariable=self.save_path_var, 
            font=self.normal_font,
            width=40
        )
        self.save_path_entry.pack(fill=tk.X)
        
        self.save_button = tk.Button(
            self.save_frame, 
            text="💾 選擇儲存位置", 
            command=self.select_save_path,
            font=self.button_font,
            bg=self.button_color,
            fg=self.button_text_color,
            activebackground=self.accent_color,
            activeforeground=self.button_text_color,
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.save_button.pack(pady=(10, 0))
        
        # 進度區
        self.progress_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        self.progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            self.progress_frame, 
            orient="horizontal", 
            length=460, 
            mode="determinate",
            variable=self.progress_var
        )
        self.progress.pack(fill=tk.X)
        
        self.progress_label = tk.Label(
            self.progress_frame, 
            text="準備就緒", 
            font=self.normal_font, 
            bg=self.bg_color, 
            fg=self.text_color
        )
        self.progress_label.pack(pady=(5, 0))
        
        # 按鈕區
        self.button_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        self.button_frame.pack(fill=tk.X)
        
        self.start_button = tk.Button(
            self.button_frame, 
            text="🚀 開始執行", 
            command=self.run_processing,
            font=self.button_font,
            bg=self.button_color,
            fg=self.button_text_color,
            activebackground=self.accent_color,
            activeforeground=self.button_text_color,
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.open_result_button = tk.Button(
            self.button_frame, 
            text="📊 開啟結果檔案", 
            command=self.open_result_file,
            font=self.button_font,
            bg="#6c757d",
            fg=self.button_text_color,
            activebackground="#5a6268",
            activeforeground=self.button_text_color,
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.open_result_button.pack(side=tk.LEFT)
        
        # 底部資訊區
        self.footer_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        self.footer_frame.pack(fill=tk.X, pady=(15, 0))
        
        current_year = datetime.now().year
        self.footer_label = tk.Label(
            self.footer_frame, 
            text=f"© {current_year} Excel 姓名統計工具", 
            font=("Microsoft JhengHei UI", 8), 
            bg=self.bg_color, 
            fg="#999999"
        )
        self.footer_label.pack(side=tk.LEFT)
        
        self.help_label = tk.Label(
            self.footer_frame, 
            text="需要幫助?", 
            font=("Microsoft JhengHei UI", 8), 
            bg=self.bg_color, 
            fg=self.accent_color,
            cursor="hand2"
        )
        self.help_label.pack(side=tk.RIGHT)
        self.help_label.bind("<Button-1>", self.show_help)
        
        # 路徑暫存
        self.file_path = None
        self.save_path = "產出報表.xlsx"
        
        # 設定 ttk 樣式
        self.style = ttk.Style()
        self.style.configure("TProgressbar", thickness=10, troughcolor="#e9ecef", background=self.accent_color)
        
        # 綁定按鍵事件
        master.bind("<Return>", lambda event: self.run_processing() if self.start_button["state"] == tk.NORMAL else None)
        
        # 設定視窗置中
        self.center_window()
    
    def center_window(self):
        """將視窗置中於螢幕"""
        self.master.update_idletasks()
        width = self.master.winfo_width()
        height = self.master.winfo_height()
        x = (self.master.winfo_screenwidth() // 2) - (width // 2)
        y = (self.master.winfo_screenheight() // 2) - (height // 2)
        self.master.geometry(f"{width}x{height}+{x}+{y}")

    def select_file(self):
        """選擇 Excel 檔案"""
        file_path = filedialog.askopenfilename(
            title="請選擇含有工作表1的 Excel 檔案",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.file_path = file_path
            self.file_path_var.set(file_path)
            self.save_button.config(state=tk.NORMAL)
            
            # 自動設定儲存檔案名稱
            file_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(file_name)[0]
            self.save_path_var.set(f"{name_without_ext}_統計結果.xlsx")
            
            # 啟用開始按鈕
            self.start_button.config(state=tk.NORMAL)

    def select_save_path(self):
        """選擇儲存位置"""
        initial_file = self.save_path_var.get()
        save_path = filedialog.asksaveasfilename(
            title="儲存報表為...",
            defaultextension=".xlsx",
            initialfile=initial_file,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if save_path:
            self.save_path = save_path
            self.save_path_var.set(save_path)

    def run_processing(self):
        """執行處理程序"""
        # 檢查檔案路徑
        if not self.file_path:
            messagebox.showwarning("警告", "請先選擇 Excel 檔案")
            return
        
        # 取得儲存路徑
        self.save_path = self.save_path_var.get()
        
        # 禁用按鈕
        self.select_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        
        # 使用 Thread 避免 GUI 卡死
        threading.Thread(target=self.process_excel).start()

    def process_excel(self):
        """處理 Excel 檔案"""
        try:
            # 更新進度
            self.update_progress(10, "讀取 Excel 檔案中...")
            time.sleep(0.3)
            
            # 讀取 Excel 檔案
            xlsx = pd.ExcelFile(self.file_path)
            
            # 檢查是否有工作表1
            if '工作表1' not in xlsx.sheet_names:
                self.master.after(0, lambda: messagebox.showerror("錯誤", "Excel 檔案中沒有「工作表1」"))
                self.reset_ui()
                return
                
            self.update_progress(30, "分析工作表資料中...")
            time.sleep(0.3)
            
            # 讀取工作表1
            df1 = xlsx.parse('工作表1')
            self.update_progress(50, "提取姓名資料中...")
            time.sleep(0.3)
            
            # 提取姓名
            data = df1.iloc[1:, 2].dropna().astype(str)
            names = data.str.extract(r'-(.+)$')[0].str.strip()
            
            # 統計姓名出現次數
            self.update_progress(70, "統計姓名出現次數中...")
            time.sleep(0.3)
            
            summary = names.value_counts().reset_index()
            summary.columns = ['姓名', '次數']
            total = summary['次數'].sum()
            summary.loc[len(summary)] = ['總計', total]
            
            # 計算百分比
            summary['百分比'] = summary['次數'].apply(lambda x: f"{(x / total * 100):.2f}%" if x != total else "100%")
            
            # 儲存結果
            self.update_progress(85, "產生報表中...")
            time.sleep(0.3)
            
            with pd.ExcelWriter(self.save_path, engine='openpyxl') as writer:
                # 寫入統計結果
                summary.to_excel(writer, sheet_name='工作表2', startrow=2, startcol=1, index=False)
                
                # 取得工作表
                ws = writer.sheets['工作表2']
                
                # 設定標題
                ws['B1'] = '姓名統計結果'
                ws['B2'] = '姓名'
                ws['C2'] = '次數'
                ws['D2'] = '百分比'
            
            self.update_progress(100, "完成！")
            time.sleep(0.5)
            
            # 顯示成功訊息
            self.master.after(0, lambda: messagebox.showinfo("完成", f"報表已成功儲存至：\n{self.save_path}"))
            
            # 啟用開啟結果按鈕
            self.master.after(0, lambda: self.open_result_button.config(state=tk.NORMAL))
            
            # 重設 UI
            self.reset_ui(keep_paths=True)
            
        except Exception as e:
            self.master.after(0, lambda: messagebox.showerror("錯誤", f"處理失敗：\n{str(e)}"))
            self.reset_ui()

    def update_progress(self, value, text):
        """更新進度條"""
        self.progress_var.set(value)
        self.master.after(0, lambda: self.progress_label.config(text=text))
        self.master.update_idletasks()

    def reset_ui(self, keep_paths=False):
        """重設 UI 狀態"""
        self.select_button.config(state=tk.NORMAL)
        self.save_button.config(state=tk.NORMAL if self.file_path else tk.DISABLED)
        self.start_button.config(state=tk.NORMAL if self.file_path else tk.DISABLED)
        self.progress_var.set(0)
        self.progress_label.config(text="準備就緒")
        
        if not keep_paths:
            self.file_path = None
            self.file_path_var.set("")
            self.save_path_var.set("產出報表.xlsx")

    def open_result_file(self):
        """開啟結果檔案"""
        if os.path.exists(self.save_path):
            try:
                os.startfile(self.save_path)
            except:
                # 如果 os.startfile 不可用 (非 Windows 系統)
                try:
                    import subprocess
                    subprocess.Popen(['xdg-open', self.save_path])  # Linux
                except:
                    try:
                        subprocess.Popen(['open', self.save_path])  # macOS
                    except:
                        messagebox.showerror("錯誤", "無法開啟檔案")
        else:
            messagebox.showerror("錯誤", f"找不到檔案：{self.save_path}")

    def show_help(self, event=None):
        """顯示說明視窗"""
        help_window = tk.Toplevel(self.master)
        help_window.title("使用說明")
        help_window.geometry("400x300")
        help_window.resizable(False, False)
        help_window.configure(bg=self.bg_color)
        
        # 設定視窗置中
        help_window.transient(self.master)
        help_window.update_idletasks()
        width = help_window.winfo_width()
        height = help_window.winfo_height()
        x = self.master.winfo_rootx() + (self.master.winfo_width() // 2) - (width // 2)
        y = self.master.winfo_rooty() + (self.master.winfo_height() // 2) - (height // 2)
        help_window.geometry(f"{width}x{height}+{x}+{y}")
        
        # 說明內容
        help_frame = tk.Frame(help_window, bg=self.bg_color, padx=20, pady=20)
        help_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = tk.Label(
            help_frame, 
            text="Excel 姓名統計工具使用說明", 
            font=self.title_font, 
            bg=self.bg_color, 
            fg=self.accent_color
        )
        title_label.pack(pady=(0, 15))
        
        help_text = """
1. 點擊「選擇檔案」按鈕，選擇含有「工作表1」的 Excel 檔案。

2. 程式會自動設定儲存位置，您也可以點擊「選擇儲存位置」按鈕來變更。

3. 點擊「開始執行」按鈕，程式會從工作表1中提取姓名並統計出現次數。

4. 處理完成後，您可以點擊「開啟結果檔案」按鈕查看統計結果。

注意：程式會從第三欄 (C 欄) 資料中提取「-」後面的姓名進行統計。
        """
        
        text_widget = tk.Text(
            help_frame, 
            wrap=tk.WORD, 
            font=self.normal_font,
            bg=self.bg_color,
            fg=self.text_color,
            relief=tk.FLAT,
            height=10
        )
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        close_button = tk.Button(
            help_frame, 
            text="關閉", 
            command=help_window.destroy,
            font=self.button_font,
            bg=self.button_color,
            fg=self.button_text_color,
            activebackground=self.accent_color,
            activeforeground=self.button_text_color,
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2"
        )
        close_button.pack(pady=(15, 0))

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernExcelToolApp(root)
    root.mainloop()
