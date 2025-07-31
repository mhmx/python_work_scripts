# Создает pdf-файл из первых страниц всех вложенных в папку pdf-файлов

import fitz
import os

input_folder =  r"C:\***"                                       # путь к папке, в конце без \
output_pdf =    input_folder + "\merged.pdf"                    # итоговый файл

pdf_writer = fitz.open()

for filename in sorted(os.listdir(input_folder)):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(input_folder, filename)
        pdf = fitz.open(pdf_path)
        if len(pdf) > 0:                                        # если есть ли страницы
            pdf_writer.insert_pdf(pdf, from_page=0, to_page=0)  # берём только первую

pdf_writer.save(output_pdf)
pdf_writer.close()

print(f"Файл {output_pdf} создан.")