import os
import pandas as pd
import requests
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from PIL import Image
import webbrowser

# --- Логика скачивания ---
def download_photos(excel_path, article_col, photo_cols, progress_callback, log_callback):
    try:
        df = pd.read_excel(excel_path, header=None)
    except Exception as e:
        log_callback(f"❌ Ошибка при чтении {excel_path}: {e}")
        return

    total_photos = sum(df[c].notna().sum() for c in photo_cols if c < len(df.columns))
    done_photos = 0

    if total_photos == 0:
        log_callback(f"⚠️ Нет ссылок в файле: {excel_path}")
        return

    base_folder = os.path.dirname(excel_path)

    for idx, row in df.iterrows():
        article_val = row[article_col] if article_col < len(row) else None
        if pd.isna(article_val):
            continue
        article = str(article_val).strip().replace(".0", "")
        if not article:
            continue

        folder = os.path.join(base_folder, article)
        os.makedirs(folder, exist_ok=True)

        for i, col in enumerate(photo_cols, start=1):
            if col >= len(row):
                continue
            url = row[col]
            if pd.notna(url):
                try:
                    r = requests.get(url, timeout=10)
                    if r.status_code == 200:
                        filename = os.path.join(folder, f"{article}_{i}.jpg")
                        with open(filename, "wb") as f:
                            f.write(r.content)
                        log_callback(f"✅ {filename}")
                    else:
                        log_callback(f"⚠️ Ошибка {r.status_code} при скачивании {url}")
                except Exception as e:
                    log_callback(f"❌ Ошибка при скачивании {url}: {e}")
                done_photos += 1
                progress_callback(done_photos, total_photos)

    log_callback(f"🎉 Готово для {excel_path}!")


# --- Конвертор изображений ---
def convert_images(folder_path, target_format, log_callback):
    if not os.path.isdir(folder_path):
        log_callback(f"❌ Папка не найдена: {folder_path}")
        return

    supported_formats = ["PNG", "JPEG", "WEBP"]
    target_format_upper = target_format.upper()
    if target_format_upper == "JPG":
        target_format_upper = "JPEG"
    if target_format_upper not in supported_formats:
        log_callback(f"❌ Неподдерживаемый формат: {target_format}")
        return

    files = os.listdir(folder_path)
    converted = 0

    for file in files:
        full_path = os.path.join(folder_path, file)
        if not os.path.isfile(full_path):
            continue
        name, ext = os.path.splitext(file)
        try:
            with Image.open(full_path) as img:
                new_file = os.path.join(folder_path, f"{name}.{target_format.lower()}")
                img.convert("RGB").save(new_file, target_format_upper)
                log_callback(f"✅ {file} → {new_file}")
                converted += 1
        except Exception as e:
            log_callback(f"⚠️ {file} нельзя конвертировать: {e}")

    log_callback(f"🎉 Конвертация завершена. Всего файлов конвертировано: {converted}")


# --- GUI ---
def browse_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)


def browse_folder(entry):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry.delete(0, tk.END)
        entry.insert(0, folder_path)


def remove_file_entry(entry, widgets):
    global file_entries
    if len(file_entries) <= 1:
        messagebox.showwarning("Внимание", "Должен быть хотя бы один файл.")
        return
    for w in widgets:
        w.destroy()
    file_entries.remove(entry)


def add_file_entry():
    global file_entries
    if len(file_entries) >= 8:
        messagebox.showwarning("Лимит", "Можно добавить максимум 8 файлов.")
        return

    row = len(file_entries)
    entry = tk.Entry(file_frame, width=70)
    entry.grid(row=row, column=0, padx=5, pady=2, sticky="ew")

    btn_browse = tk.Button(file_frame, text="Обзор", command=lambda e=entry: browse_file(e))
    btn_browse.grid(row=row, column=1, padx=2)

    btn_remove = tk.Button(file_frame, text="–")
    btn_remove.config(command=lambda e=entry, w=[entry, btn_browse, btn_remove]: remove_file_entry(e, w))
    btn_remove.grid(row=row, column=2, padx=2)

    file_entries.append(entry)


def start_download():
    excel_paths = [e.get().strip() for e in file_entries if e.get().strip()]
    if not excel_paths:
        messagebox.showerror("Ошибка", "Добавьте хотя бы один Excel файл.")
        return

    try:
        article_col = int(entry_article_col.get())
        photo_cols = [int(c.strip()) for c in entry_photo_cols.get().split(",")]
    except:
        messagebox.showerror("Ошибка", "Введите корректные номера колонок.")
        return

    text_log.delete(1.0, tk.END)

    def run():
        total_files = len(excel_paths)
        for idx, path in enumerate(excel_paths, start=1):
            log_callback(f"📂 Обработка файла ({idx}/{total_files}): {path}")
            download_photos(path, article_col, photo_cols, progress_callback, log_callback)

    threading.Thread(target=run, daemon=True).start()

def delete_files_of_format():
    folder_path = entry_convert_folder.get().strip()
    target_format = combo_format.get()
    if not folder_path or not target_format:
        messagebox.showerror("Ошибка", "Укажите папку и формат для удаления.")
        return
    if not os.path.isdir(folder_path):
        messagebox.showerror("Ошибка", f"Папка не найдена: {folder_path}")
        return

    # Подтверждение удаления
    if not messagebox.askyesno("Подтвердите", f"Удалить все файлы формата .{target_format} в папке?"):
        return

    threading.Thread(target=lambda: _delete_files_worker(folder_path, target_format), daemon=True).start()

def _delete_files_worker(folder_path, target_format):
    deleted = 0
    for file in os.listdir(folder_path):
        if file.lower().endswith(f".{target_format.lower()}"):
            try:
                os.remove(os.path.join(folder_path, file))
                log_callback(f"🗑️ Удалено: {file}")
                deleted += 1
            except Exception as e:
                log_callback(f"❌ Не удалось удалить {file}: {e}")
    log_callback(f"🎉 Удаление завершено. Всего удалено: {deleted} файлов.")    


def start_conversion():
    folder_path = entry_convert_folder.get().strip()
    target_format = combo_format.get()
    if not folder_path or not target_format:
        messagebox.showerror("Ошибка", "Укажите папку и формат для конвертации.")
        return
    threading.Thread(target=lambda: convert_images(folder_path, target_format, log_callback), daemon=True).start()


def log_callback(msg):
    text_log.insert(tk.END, msg + "\n")
    text_log.see(tk.END)


def progress_callback(done, total):
    percent = int(done / total * 100)
    progress_bar["value"] = percent
    lbl_progress.config(text=f"{percent}% ({done}/{total})")
    root.update_idletasks()


def open_link(url):
    webbrowser.open(url)


# --- GUI Window ---
root = tk.Tk()
root.title("Photo Downloader & Converter v4.0")
root.geometry("900x700")
root.minsize(800, 600)

# Инструкция
lbl_instr = tk.Label(root, text=(
    "Инструкция:\n"
    "1. Добавьте Excel-файлы с данными.\n"
    "2. Укажите колонку с артикулами и колонки со ссылками.\n"
    "3. Нажмите 'Начать скачивание' для сохранения фото.\n"
    "4. В секции 'Конвертор' укажите папку с фото и формат для конвертации."
), justify="left", wraplength=850, fg="blue")
lbl_instr.pack(pady=10)

# Файлы Excel
file_frame = tk.Frame(root)
file_frame.pack(pady=10, fill="x")
file_frame.columnconfigure(0, weight=1)
file_entries = []
add_file_entry()
btn_add = tk.Button(root, text="+ Добавить файл", command=add_file_entry)
btn_add.pack(pady=5)

# Настройки колонок
frame_cols = tk.Frame(root)
frame_cols.pack(pady=5, fill="x")
frame_cols.columnconfigure(1, weight=1)
frame_cols.columnconfigure(3, weight=1)

tk.Label(frame_cols, text="Колонка с артикулами (номер):").grid(row=0, column=0, sticky="e")
entry_article_col = tk.Entry(frame_cols, width=5)
entry_article_col.insert(0, "1")
entry_article_col.grid(row=0, column=1, padx=5, sticky="ew")

tk.Label(frame_cols, text="Колонки со ссылками (через запятую):").grid(row=0, column=2, sticky="e")
entry_photo_cols = tk.Entry(frame_cols, width=15)
entry_photo_cols.insert(0, "2,3,4,5,6")
entry_photo_cols.grid(row=0, column=3, padx=5, sticky="ew")

# Кнопка скачивания
btn_start = tk.Button(root, text="Начать скачивание", command=start_download, bg="green", fg="white")
btn_start.pack(pady=10)

# Конвертор изображений
frame_convert = tk.LabelFrame(root, text="Конвертор формата фотографий")
frame_convert.pack(pady=10, fill="x", padx=5)
frame_convert.columnconfigure(1, weight=1)

tk.Label(frame_convert, text="Папка с изображениями:").grid(row=0, column=0, sticky="e")
entry_convert_folder = tk.Entry(frame_convert, width=70)
entry_convert_folder.grid(row=0, column=1, padx=5, sticky="ew")
btn_browse_folder = tk.Button(frame_convert, text="Обзор", command=lambda: browse_folder(entry_convert_folder))
btn_browse_folder.grid(row=0, column=2, padx=2)

tk.Label(frame_convert, text="Формат:").grid(row=1, column=0, sticky="e")
combo_format = ttk.Combobox(frame_convert, values=["png", "jpg", "webp"], width=10)
combo_format.current(0)
combo_format.grid(row=1, column=1, sticky="w", padx=5)

btn_convert = tk.Button(frame_convert, text="Конвертировать", command=start_conversion, bg="orange", fg="white")
btn_convert.grid(row=2, column=1, pady=5, sticky="w")

btn_delete = tk.Button(frame_convert, text="Удалить все файлы формата", command=delete_files_of_format, bg="red", fg="white")
btn_delete.grid(row=2, column=2, padx=5, sticky="w")

# Прогресс
progress_bar = ttk.Progressbar(root, length=700)
progress_bar.pack(pady=5, fill="x", padx=5)
lbl_progress = tk.Label(root, text="0%")
lbl_progress.pack()

# Лог
text_log = scrolledtext.ScrolledText(root, width=110, height=20)
text_log.pack(pady=10, fill="both", expand=True)

# Подпись внизу
frame_footer = tk.Frame(root)
frame_footer.pack(fill="x", pady=5, padx=10)
frame_footer.columnconfigure(0, weight=1)
frame_footer.columnconfigure(1, weight=1)

lbl_left = tk.Label(frame_footer, text="Программа от Three_Per_Cento", fg="blue", cursor="hand2")
lbl_left.grid(row=0, column=0, sticky="w")
lbl_left.bind("<Button-1>", lambda e: open_link("https://github.com/ThreePerCento"))

lbl_right = tk.Label(frame_footer, text="GitHub program", fg="blue", cursor="hand2")
lbl_right.grid(row=0, column=1, sticky="e")
lbl_right.bind("<Button-1>", lambda e: open_link("https://github.com/ThreePerCento/Photo_downloader_from_links_to_Excel/releases"))

root.mainloop()
