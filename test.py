import os
from PyPDF2 import PdfReader

import os
import win32com.client

def docx_to_pdf(docx_path, pdf_path="output.pdf"):
    # Проверяем существование входного файла
    if not os.path.exists(docx_path):
        print(f"Ошибка: Файл {docx_path} не найден")
        return None
    
    # Абсолютные пути
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path)
    
    print(f"Конвертация {docx_path} в {pdf_path}")
    
    try:
        # Создаём объект Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Скрываем Word для скорости
        
        # Открываем документ
        doc = word.Documents.Open(docx_path)
        
        # Сохраняем как PDF (17 — это wdFormatPDF)
        doc.SaveAs(pdf_path, FileFormat=17)
        
        # Закрываем документ и Word
        doc.Close()
        word.Quit()
        
        print(f"Файл успешно сохранён как {pdf_path}")
        return pdf_path
    
    except Exception as e:
        print(f"Ошибка при конвертации: {e}")
        return None

def get_page_count_via_pdf(pdf_path):
    # Открываем PDF и считаем страницы
    pdf_file = PdfReader(pdf_path)
    page_count = len(pdf_file.pages)
    
    # Опционально: удаляем временный PDF-файл
    os.remove(pdf_path)
    
    return page_count

# Пример использования
if __name__ == "__main__":
    # Тест для Word
    word_file = os.path.join(os.getcwd(), "Documents\sppd\КПП", "КПП_1.docx")
    pdf_path = docx_to_pdf(word_file)
    word_pages = get_page_count_via_pdf(pdf_path)
    print(f"Количество страниц в {word_file}: {word_pages}")
