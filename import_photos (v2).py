import os
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

def download_file(url, filename):
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200:
            with open(filename, "wb") as f:
                f.write(r.content)
            return True
        return False
    except:
        return False

def download_images_from_excel(file_path):
    print(f"\n📂 Обработка файла: {file_path}")
    df = pd.read_excel(file_path, header=None)

    # базовая папка = папка, где лежит Excel
    base_folder = os.path.dirname(file_path)

    tasks = []
    for idx, row in df.iterrows():
    article = row[1]  # колонка B (артикул)
    if pd.isna(article):
        continue  # пропускаем пустые строки

    # Преобразуем к строке без лишнего ".0" и пробелов
    article = str(article).strip()
    if article.endswith(".0"):
        article = article[:-2]

    folder = os.path.join(base_folder, article)
    os.makedirs(folder, exist_ok=True)

        for i, col in enumerate([2, 3, 4, 5, 6], start=1):
            url = row[col]
            if pd.notna(url):
                filename = os.path.join(folder, f"{article}_{i}.jpg")
                tasks.append((url, filename))

    total = len(tasks)
    print(f"🔎 Найдено {total} фотографий для скачивания.")

    completed = 0
    with ThreadPoolExecutor(max_workers=8) as executor:
        future_to_task = {executor.submit(download_file, url, filename): (url, filename) for url, filename in tasks}

        for future in as_completed(future_to_task):
            url, filename = future_to_task[future]
            completed += 1
            percent = (completed / total) * 100
            try:
                result = future.result()
                if result:
                    print(f"[{percent:.1f}%] ✅ Скачано: {filename}")
                else:
                    print(f"[{percent:.1f}%] ⚠️ Ошибка скачивания {url}")
            except Exception as e:
                print(f"[{percent:.1f}%] ❌ Ошибка {url}: {e}")

    print(f"🎉 Готово! Скачано {completed} из {total} фотографий.")

if __name__ == "__main__":
    excel_files = [
        r"C:\Users\3prokent\Documents\Python\кокошники\data(кокошники).xlsx",
        r"C:\Users\3prokent\Documents\Python\хохлома\data(хохлома).xlsx",
    ]

    for file in excel_files:
        download_images_from_excel(file)