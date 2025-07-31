import pandas as pd
import os
from PyPDF2 import PdfReader
import zipfile
import rarfile
import py7zr
import win32com.client as win32
from tqdm import tqdm

INPUT_FILE = "Дозагрузить.xlsx"
OUTPUT_FILE = "files_with_print_pages.xlsx"

print("Загрузка списка файлов...")
df = pd.read_excel(INPUT_FILE)

print("Запуск Excel и Word...")
excel_app = win32.gencache.EnsureDispatch('Excel.Application')
excel_app.Visible = False
word_app = win32.gencache.EnsureDispatch('Word.Application')
word_app.Visible = False

def analyze_file(path):
    try:
        if not os.path.exists(path):
            return "Файл не найден"

        ext = os.path.splitext(path)[1].lower()

        if ext == ".pdf":
            with open(path, 'rb') as f:
                reader = PdfReader(f)
                return len(reader.pages)

        elif ext in [".xlsx", ".xls", ".xlsm"]:
            wb = excel_app.Workbooks.Open(path, ReadOnly=True)
            total_pages = 0
            for sheet in wb.Sheets:
                try:
                    sheet.Activate()
                    pages = sheet.PageSetup.Pages.Count
                except:
                    pages = excel_app.ExecuteExcel4Macro(f'GET.DOCUMENT(50, "{sheet.Name}")')
                total_pages += pages
            wb.Close(SaveChanges=False)
            return total_pages

        elif ext in [".doc", ".docx"]:
            doc = word_app.Documents.Open(path, ReadOnly=True)
            doc.Repaginate()
            pages = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
            doc.Close(SaveChanges=False)
            return pages

        elif ext == ".zip":
            with zipfile.ZipFile(path, 'r') as zf:
                return f"{len(zf.namelist())} файлов в архиве"

        elif ext == ".rar":
            with rarfile.RarFile(path, 'r') as rf:
                return f"{len(rf.namelist())} файлов в архиве"

        elif ext == ".7z":
            with py7zr.SevenZipFile(path, mode='r') as z:
                return f"{len(z.getnames())} файлов в архиве"

        else:
            return "Не поддерживается"

    except Exception as e:
        return f"Ошибка: {e}"

# Оборачиваем для tqdm и вывода ошибок
def analyze_with_progress(path):
    result = analyze_file(path)
    if isinstance(result, str) and (
        "не найден" in result.lower() or
        "не поддерживается" in result.lower() or
        "ошибка" in result.lower()
    ):
        print(f"\n[!] {path} — {result}")
    return result

print("Обработка файлов...")
tqdm.pandas(desc="Обработка")
df["Страниц (печать)"] = df.iloc[:, 0].progress_apply(analyze_with_progress)

print("Закрытие Excel и Word...")
excel_app.Quit()
word_app.Quit()

def save_with_fallback(df, base_filename):
    name, ext = os.path.splitext(base_filename)
    counter = 1
    while True:
        try:
            df.to_excel(base_filename, index=False)
            print(f"Результат сохранён в файл: {base_filename}")
            break
        except PermissionError:
            new_name = f"{name}_{counter}{ext}"
            print(f"[!] Невозможно сохранить файл (занят?). Пытаюсь: {new_name}")
            base_filename = new_name
            counter += 1

print("Сохраняем результат...")
save_with_fallback(df, OUTPUT_FILE)
print("Готово!")
