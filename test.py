import tkinter as tk
from tkinter import filedialog
from PyPDF2 import PdfReader, PdfWriter
import os
import fitz  # PyMuPDF

# def compress_pdf(input_path, output_path):
#     """
#     Сжимает PDF файл с использованием встроенной компрессии PyPDF2
#     """
#     try:
#         # Читаем исходный PDF
#         reader = PdfReader(input_path)
#         writer = PdfWriter()

#         # Копируем все страницы с компрессией
#         for page in reader.pages:
#             # Сжимаем содержимое страницы (без указания уровня)
#             page.compress_content_streams()
#             writer.add_page(page)

#         # Включаем компрессию для всего документа
#         writer.add_metadata(reader.metadata)  # Сохраняем метаданные если есть
        
#         # Сохраняем сжатый PDF
#         with open(output_path, 'wb') as output_file:
#             writer.write(output_file)
        
#         # Выводим информацию о сжатии
#         original_size = os.path.getsize(input_path) / 1024  # в KB
#         compressed_size = os.path.getsize(output_path) / 1024  # в KB
#         print(f"Оригинальный размер: {original_size:.2f} KB")
#         print(f"Сжатый размер: {compressed_size:.2f} KB")
#         print(f"Сжатие: {((original_size - compressed_size) / original_size * 100):.2f}%")
        
#         return True
#     except Exception as e:
#         print(f"Ошибка при сжатии: {str(e)}")
#         return False

def compress_pdf(input_path, output_path, quality=2):
    try:
        doc = fitz.open(input_path)
        for page in doc:
            for img in page.get_images(full=True):
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n > 3:  # Конвертация в RGB, если есть альфа-канал
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                page.insert_image(page.rect, pixmap=pix)
        doc.save(output_path, garbage=4, deflate=True)
        return True
    except Exception as e:
        print(f"Ошибка при сжатии: {str(e)}")
        return False

def select_and_compress_pdf():
    # Создаем окно tkinter
    root = tk.Tk()
    root.withdraw()  # Скрываем основное окно

    # Открываем диалог выбора файла
    input_path = filedialog.askopenfilename(
        title="Выберите PDF файл",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not input_path:
        print("Файл не выбран")
        return

    # Получаем путь для выходного файла
    output_path = input_path.replace('.pdf', '_compressed.pdf')

    print("Сжатие файла...")
    success = compress_pdf(input_path, output_path)
    
    if success:
        print(f"Файл успешно сжат и сохранен как: {output_path}")
    else:
        print("Не удалось сжать файл")

if __name__ == "__main__":
    select_and_compress_pdf()