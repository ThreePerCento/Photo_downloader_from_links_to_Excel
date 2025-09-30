import os
import pandas as pd
import requests
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

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


# --- GUI ---
def browse_file(entry):
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)


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
    entry.grid(row=row, column=0, padx=5, pady=2)

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
            download_photos(
                path,
                article_col,
                photo_cols,
                progress_callback,
                log_callback
            )

    threading.Thread(target=run, daemon=True).start()


def log_callback(msg):
    text_log.insert(tk.END, msg + "\n")
    text_log.see(tk.END)


def progress_callback(done, total):
    percent = int(done / total * 100)
    progress_bar["value"] = percent
    lbl_progress.config(text=f"{percent}% ({done}/{total})")
    root.update_idletasks()


# --- GUI Window ---
root = tk.Tk()
root.title("Скачивание фотографий из Excel")
root.geometry("800x600")

# Инструкция
lbl_instr = tk.Label(
    root,
    text=(
        "Инструкция:\n"
        "1. Добавьте от 1 до 8 Excel-файлов с данными (кнопка +).\n"
        "2. В колонке с артикулами должны быть уникальные коды товаров.\n"
        "3. В колонках со ссылками — URL картинок (через запятую можно указать несколько колонок).\n"
        "4. В каждой папке (по имени артикула) сохранятся все фото.\n"
        "5. Прогресс будет показан ниже."
    ),
    justify="left",
    wraplength=750,
    fg="blue"
)
lbl_instr.pack(pady=10)

# Файлы
file_frame = tk.Frame(root)
file_frame.pack(pady=10)
file_entries = []

add_file_entry()

btn_add = tk.Button(root, text="+ Добавить файл", command=add_file_entry)
btn_add.pack(pady=5)

# Настройки колонок
frame_cols = tk.Frame(root)
frame_cols.pack(pady=10)

tk.Label(frame_cols, text="Колонка с артикулами (номер):").grid(row=0, column=0, sticky="e")
entry_article_col = tk.Entry(frame_cols, width=5)
entry_article_col.insert(0, "1")  # колонка B = индекс 1
entry_article_col.grid(row=0, column=1, padx=5)

tk.Label(frame_cols, text="Колонки со ссылками (через запятую):").grid(row=0, column=2, sticky="e")
entry_photo_cols = tk.Entry(frame_cols, width=15)
entry_photo_cols.insert(0, "2,3,4,5,6")  # по умолчанию C–G
entry_photo_cols.grid(row=0, column=3, padx=5)

# Кнопка запуска
btn_start = tk.Button(root, text="Начать скачивание", command=start_download, bg="green", fg="white")
btn_start.pack(pady=10)

# Прогресс
progress_bar = ttk.Progressbar(root, length=600)
progress_bar.pack(pady=5)
lbl_progress = tk.Label(root, text="0%")
lbl_progress.pack()

# Лог
text_log = scrolledtext.ScrolledText(root, width=95, height=15)
text_log.pack(pady=10)

root.mainloop()

