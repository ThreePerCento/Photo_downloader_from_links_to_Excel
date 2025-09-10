import os
import pandas as pd
import requests

# Загружаем Excel (без заголовков!)
df = pd.read_excel("data.xlsx", header=None)

# Перебираем строки
for idx, row in df.iterrows():
    article = str(row[1])  # колонка B (артикул)
    if pd.isna(article):
        continue  # пропускаем пустые строки
    
    # создаем папку для артикула
    folder = os.path.join("матрешки", article)
    os.makedirs(folder, exist_ok=True)
    
    # перебираем ссылки из C–G
    for i, col in enumerate([2, 3, 4, 5, 6], start=1):
        url = row[col]
        if pd.notna(url):
            try:
                r = requests.get(url, timeout=10)
                if r.status_code == 200:
                    filename = os.path.join(folder, f"{article}_{i}.jpg")
                    with open(filename, "wb") as f:
                        f.write(r.content)
                    print(f"✅ Скачано: {filename}")
                else:
                    print(f"⚠️ Ошибка {r.status_code} при скачивании {url}")
            except Exception as e:
                print(f"❌ Ошибка при скачивании {url}: {e}")

print("🎉 Готово! Все фото скачаны.")