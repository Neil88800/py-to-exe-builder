import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import os

TARGET_SIZE_KB = 200

def compress_image(file_path):
    img = Image.open(file_path)
    # PNG、RGBA 轉為 JPEG RGB
    if img.mode in ("RGBA", "LA"):
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        img = background
    else:
        img = img.convert("RGB")

    output_path = os.path.splitext(file_path)[0] + "_compressed.jpg"
    quality = 95

    while quality > 10:
        img.save(output_path, "JPEG", quality=quality, optimize=True)
        if os.path.getsize(output_path) <= TARGET_SIZE_KB * 1024:
            return output_path
        quality -= 5

    return output_path

def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.webp")])
    if file_path:
        output = compress_image(file_path)
        messagebox.showinfo("完成", f"壓縮完成，輸出檔案：\n{output}")

root = tk.Tk()
root.title("圖片壓縮工具")
root.geometry("300x150")

btn = tk.Button(root, text="選擇圖片並壓縮", command=choose_file)
btn.pack(expand=True)

root.mainloop()
