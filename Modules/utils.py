import openpyxl
import os
from datetime import datetime
from tkinter import filedialog, messagebox

def format_date(date_str):
    months = {
        "01": "января", "02": "февраля", "03": "марта", "04": "апреля",
        "05": "мая", "06": "июня", "07": "июля", "08": "августа",
        "09": "сентября", "10": "октября", "11": "ноября", "12": "декабря"
    }
    if '.' in date_str:
        day, month, year = date_str.split(".")
    return f"«{day}» {months[month]} {year}"

def get_cell_value(ws, cell_row, cell_column):
    cell = f"{cell_column}{cell_row}"
    value = ws[cell].value
    
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    
    return str(value) if value is not None else ""

def split_values(value):
    if value == "":
        return []
    if not isinstance(value, str):
        return [str(value)]
    
    values = [v.strip() for v in value.split(';') if v.strip()]
    return values

def clear_ui(root, menu_button, menu):
    for widget in root.winfo_children():
        if widget != menu_button and widget != menu:
            widget.destroy()

def select_excel_file(project_name=None):
    if project_name:
        file_path = os.path.join(os.getcwd(), "Excel_files", f"{project_name}.xlsx")
        if not os.path.exists(file_path):
            messagebox.showerror("Ошибка", f"Файл {file_path} не найден!")
            return None, None, None
    else:
        file_path = filedialog.askopenfilename(
            title="Выберите файл Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=os.path.join(os.getcwd(), "Excel_files")
        )
    if file_path:
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            excel_file_name = os.path.splitext(os.path.basename(file_path))[0]
            return wb, excel_file_name, file_path
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл: {str(e)}")
            return None, None, None
    else:
        if not project_name:  # Показываем предупреждение только если не было автоматической загрузки
            messagebox.showwarning("Предупреждение", "Файл не выбран!")
        return None, None, None
