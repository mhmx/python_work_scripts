import os
from openpyxl import Workbook

def clean_path(path):
    return path.strip().strip('"').strip("'")

def get_available_filename(base_path):
    """Возвращает доступное имя файла, если базовое занято."""
    if not os.path.exists(base_path):
        return base_path

    name, ext = os.path.splitext(base_path)
    counter = 1
    while True:
        new_path = f"{name}_({counter}){ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1

def list_files_to_excel(directory_path, output_excel_path='file_list.xlsx'):
    directory_path = clean_path(directory_path)

    if not os.path.isdir(directory_path):
        print("Указанный путь не является директорией.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Файлы"

    # Заголовки
    ws.append(["Путь к файлу", "Имя файла"])

    for root, dirs, files in os.walk(directory_path):
        for file in files:
            full_path = os.path.abspath(os.path.join(root, file))
            ws.append([full_path, file])

    output_excel_path = clean_path(output_excel_path)
    output_excel_path = get_available_filename(output_excel_path)

    try:
        wb.save(output_excel_path)
        print(f"Список файлов сохранён в: {output_excel_path}")
    except PermissionError:
        print(f"Ошибка доступа к файлу: {output_excel_path}")
        print("Проверь, не открыт ли он в Excel и попробуй снова.")
    except Exception as e:
        print(f"Ошибка при сохранении: {e}")

if __name__ == "__main__":
    folder = input("Введите путь к директории: ").strip()
    list_files_to_excel(folder)
