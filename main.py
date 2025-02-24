import tkinter as tk
from tkinter import ttk, Menu
from Modules.create_project import CreateProject
from Modules.create_acts import CreateActs
from Modules.create_registry import CreateRegistry

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Создание документов")
        self.root.geometry("800x600")
        self.root.grid_rowconfigure(11, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        # Кнопка "Меню"
        self.menu_button = ttk.Button(self.root, text="Меню", command=self.show_menu)
        self.menu_button.grid(row=0, column=0, padx=5, pady=5, sticky="nw")

        # Создание выпадающего меню
        self.menu = Menu(self.root, tearoff=0)
        self.menu.add_command(label="Создать проект", command=self.load_create_project)
        self.menu.add_command(label="Создать акты", command=self.load_create_acts)
        self.menu.add_command(label="Создать реестр", command=self.load_create_registry)

        # Инициализация модулей
        self.create_project = CreateProject(self.root, self.menu_button, self.menu)
        self.create_acts = CreateActs(self.root, self.menu_button, self.menu)
        self.create_registry = CreateRegistry(self.root, self.menu_button, self.menu)

    def show_menu(self):
        # Отображаем меню под кнопкой
        try:
            x = self.menu_button.winfo_rootx()
            y = self.menu_button.winfo_rooty() + self.menu_button.winfo_height()
            self.menu.tk_popup(x, y)
        except Exception as e:
            print(f"Ошибка при отображении меню: {e}")

    def load_create_project(self):
        project_name = self.create_project.load_ui()
        if project_name:  # Только если проект успешно создан
            self.create_acts.load_ui(project_name)

    def load_create_acts(self):
        self.create_acts.load_ui()

    def load_create_registry(self):
        self.create_registry.load_ui()

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()