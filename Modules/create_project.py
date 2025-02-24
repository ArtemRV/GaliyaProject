import openpyxl
import os
import shutil
from tkinter import messagebox, simpledialog
from Modules.utils import clear_ui

class CreateProject:
    def __init__(self, root, menu_button, menu):
        self.root = root
        self.menu_button = menu_button
        self.menu = menu
        self.wb = None
        self.excel_file_name = None
        self.excel_file_path = None
        self.entries = {}

    def load_ui(self):
        project_name = simpledialog.askstring("Создать проект", "Введите название проекта:")
        if project_name:
            template_path = os.path.join(os.getcwd(), "Templates/Excel_templates", "Template.xlsx")
            new_file_path = os.path.join(os.getcwd(), "Excel_files", f"{project_name}.xlsx")
            
            try:
                if not os.path.exists(template_path):
                    messagebox.showerror("Ошибка", "Файл Template.xlsx не найден в папке Templates!")
                    return None
                
                shutil.copyfile(template_path, new_file_path)
                self.wb = openpyxl.load_workbook(new_file_path)
                self.excel_file_name = project_name
                self.excel_file_path = new_file_path
                clear_ui(self.root, self.menu_button, self.menu)
                messagebox.showinfo("Успех", f"Проект '{project_name}' создан и загружен: {new_file_path}")
                return project_name
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось создать проект: {str(e)}")
                return None        
        else:
            messagebox.showwarning("Предупреждение", "Название проекта не введено!")
            return None
