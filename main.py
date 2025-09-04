import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
import os, zipfile

TARGET_SIZE_KB = 200

def compress_image(file_path, output_dir):
    img = Image.open(file_path)
    if img.mode in ("RGBA", "LA"):
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        img = background
    else:
        img = img.convert("RGB")

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}_compressed.jpg")
    quality = 95
    while quality > 10:
        img.save(output_path, "JPEG", quality=quality, optimize=True)
        if os.path.getsize(output_path) <= TARGET_SIZE_KB * 1024:
            return output_path
        quality -= 5
    return output_path

def choose_files():
    file_paths = filedialog.askopenfilenames(
        title="選擇圖片",
        filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.webp")]
    )
    if not file_paths:
        return

    output_dir = filedialog.askdirectory(title="選擇輸出資料夾")
    if not output_dir:
        return

    zip_filename = os.path.join(output_dir, "compressed_images.zip")
    with zipfile.ZipFile(zip_filename, "w") as zipf:
        for i, f in enumerate(file_paths, start=1):
            compressed_file = compress_image(f, output_dir)
            zipf.write(compressed_file)
            os.remove(compressed_file)
            progress_var.set(f"正在壓縮: {i}/{len(file_paths)}")
            root.update_idletasks()

    messagebox.showinfo("完成", f"已壓縮 {len(file_paths)} 張圖片並打包成\n{zip_filename}")
    progress_var.set("完成 ✅")

# --- GUI 美化 ---
root = tk.Tk()
root.title("圖片批次壓縮工具")
root.geometry("450x250")
root.resizable(False, False)

# 標題
title_label = ttk.Label(root, text="圖片批次壓縮工具", font=("Helvetica", 16, "bold"))
title_label.pack(pady=10)

# 說明文字
desc_label = ttk.Label(root, text="選擇圖片 → 壓縮到 ≤200KB → 打包成 ZIP", font=("Arial", 11))
desc_label.pack(pady=5)

# 進度文字
progress_var = tk.StringVar()
progress_var.set("等待操作...")
progress_label = ttk.Label(root, textvariable=progress_var, font=("Arial", 11), foreground="green")
progress_label.pack(pady=5)

# 按鈕
style = ttk.Style()
style.configure("TButton", font=("Arial", 12), padding=6)
btn = ttk.Button(root, text="選擇圖片並批次壓縮", command=choose_files)
btn.pack(pady=15)

# 開發者資訊
dev_label = ttk.Label(root, text="開發者 Neil", font=("Arial", 9, "italic"), foreground="gray")
dev_label.pack(side="bottom", pady=10)

root.mainloop()
