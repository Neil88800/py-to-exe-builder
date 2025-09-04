import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import os, zipfile

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

    output_path = os.path.splitext(os.path.basename(file_path))[0] + "_compressed.jpg"
    quality = 95

    while quality > 10:
        img.save(output_path, "JPEG", quality=quality, optimize=True)
        if os.path.getsize(output_path) <= TARGET_SIZE_KB * 1024:
            return output_path
        quality -= 5

    return output_path

def choose_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.webp")])
    if not file_paths:
        return

    zip_filename = "compressed_images.zip"
    with zipfile.ZipFile(zip_filename, "w") as zipf:
        for i, f in enumerate(file_paths, start=1):
            compressed_file = compress_image(f)
            zipf.write(compressed_file)
            os.remove(compressed_file)  # 壓縮完成後刪掉單張壓縮檔
            print(f"[{i}/{len(file_paths)}] 已壓縮: {compressed_file}")

    messagebox.showinfo("完成", f"已壓縮 {len(file_paths)} 張圖片並打包成\n{zip_filename}")

root = tk.Tk()
root.title("圖片批次壓縮工具")
root.geometry("350x150")

btn = tk.Button(root, text="選擇圖片並批次壓縮", command=choose_files)
btn.pack(expand=True)

root.mainloop()
