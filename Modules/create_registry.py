import tkinter as tk
from tkinter import messagebox, ttk
from Modules.utils import clear_ui, select_excel_file, get_cell_value

main_data = [
    ({'cell_row': '1'}, {'name': '_OBJECT-NAME_', 'cell_row': '1', 'cell_column': 'B'}),
    ({'cell_row': '2'}, {'name': '_CONTRACTOR-REPRESENTATIVE_', 'cell_row': '2', 'cell_column': 'B'}),
    ({'cell_row': '3'}, {'name': '_CONTRACTOR-REPRESENTATIVE-NAME_', 'cell_row': '3', 'cell_column': 'B'}),
    ({'cell_row': '4'}, {'name': '_TECHNICAL-SUPERVISION-REPRESENTATIVE_', 'cell_row': '4', 'cell_column': 'B'}),
    ({'cell_row': '5'}, {'name': '_TECHNICAL-SUPERVISION-REPRESENTATIVE-NAME_', 'cell_row': '5', 'cell_column': 'B'}),
    ({'cell_row': '6'}, {'name': '_DESIGN-ORGANIZATION-REPRESENTATIVE_', 'cell_row': '6', 'cell_column': 'B'}),
    ({'cell_row': '7'}, {'name': '_DESIGN-ORGANIZATION-REPRESENTATIVE-NAME_', 'cell_row': '7', 'cell_column': 'B'}),
    ({'cell_row': '8'}, {'name': '_ADDITIONAL-REPRESENTATIVES_', 'cell_row': '8', 'cell_column': 'B'}),
    ({'cell_row': '9'}, {'name': '_ADDITIONAL-REPRESENTATIVES-NAME_', 'cell_row': '9', 'cell_column': 'B'}),
    ({'cell_row': '10'}, {'name': '_GENERAL-CONTRACTOR_', 'cell_row': '10', 'cell_column': 'B'})
]

class CreateRegistry:
    def __init__(self, root, menu_button, menu):
        self.root = root
        self.menu_button = menu_button
        self.menu = menu
        self.wb = None
        self.excel_file_name = None
        self.excel_file_path = None
        self.entries = {}
        self.sheet_vars = {}
        self.all_var = None

    def load_ui(self, project_name=None):
        wb, excel_file_name, excel_file_path = select_excel_file(project_name)

        if wb:
            self.wb = wb
            self.excel_file_name = excel_file_name
            self.excel_file_path = excel_file_path
            self.show_acts_ui()

    def show_acts_ui(self):
        clear_ui(self.root, self.menu_button, self.menu)
        main_sheet = self.wb['Main data']
        
        self.entries.clear()
        for i, (label_data, value_data) in enumerate(main_data, start=1):
            label_text = get_cell_value(main_sheet, label_data['cell_row'], 'A')
            value = get_cell_value(main_sheet, value_data['cell_row'], 'B')
            
            tk.Label(self.root, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky="e")
            self.entries[value_data['name']] = tk.Entry(self.root, width=100)
            self.entries[value_data['name']].grid(row=i, column=1, padx=5, pady=5)
            self.entries[value_data['name']].insert(0, value)

        frame = tk.Frame(self.root)
        frame.grid(row=11, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        tk.Label(scrollable_frame, text="Выберите листы:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sheet_vars.clear()
        for i, sheet_name in enumerate(self.wb.sheetnames, start=1):
            if sheet_name not in {'Main data', 'Contents'}:
                var = tk.BooleanVar()
                tk.Checkbutton(scrollable_frame, text=sheet_name, variable=var).grid(row=i, column=0, padx=5, pady=2, sticky="w")
                self.sheet_vars[sheet_name] = var

        canvas.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")

        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        self.all_var = tk.BooleanVar()
        tk.Checkbutton(frame, text="Выбрать все", variable=self.all_var, command=self.toggle_all_sheets).grid(row=0, column=0, padx=5, pady=5, sticky="w")

        generate_button = tk.Button(self.root, text="Создать акты", command=self.generate_document)
        generate_button.grid(row=12, column=0, columnspan=2, pady=10)        
