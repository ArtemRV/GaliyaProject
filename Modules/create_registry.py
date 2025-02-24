import tkinter as tk
from Modules.utils import clear_ui

class CreateRegistry:
    def __init__(self, root, menu_button, menu):
        self.root = root
        self.menu_button = menu_button
        self.menu = menu  # Добавляем menu для исключения в clear_ui

    def load_ui(self):
        clear_ui(self.root, self.menu_button, self.menu)
        tk.Label(self.root, text="Функция создания реестра пока не реализована").grid(row=1, column=0, columnspan=2, pady=20)
