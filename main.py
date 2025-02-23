import openpyxl
from docx import Document
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

def format_date(date_str):
    months = {
        "01": "января", "02": "февраля", "03": "марта", "04": "апреля",
        "05": "мая", "06": "июня", "07": "июля", "08": "августа",
        "09": "сентября", "10": "октября", "11": "ноября", "12": "декабря"
    }
    
    day, month, year = date_str.split(".")
    return f"«{day}» {months[month]} {year}"

# Read data from a specific cell
def get_cell_value(ws, cell_row, cell_column):
    cell = f"{cell_column}{cell_row}"
    return str(ws[cell].value )

# Load source and destination files
excel_file = "template.xlsx"
word_file = "template.docx"

main_data = [
    {'name': '_OBJECTNAME_', 'cell_column': 'B', 'cell_row': '1'},
    {'name': '_SUBOBJECT-NAME_', 'cell_column': 'B', 'cell_row': '2'}
]

subobject_data = [
    {'name': 'ACT_NUMBER', 'cell_column': 'A'},
    {'name': 'EXECUTION_DATE_MONTH', 'cell_column': 'B'},
    {'name': 'EXECUTION_DATE', 'cell_column': 'B'},
    {'name': 'WORK_NAMING', 'cell_column': 'C'},
    {'name': 'ALBUM_NAME', 'cell_column': 'D'},
    {'name': 'PAGE', 'cell_column': 'E'},
    {'name': 'MATERIALS', 'cell_column': 'F'},
    {'name': 'EXECUTIVE_DIAGRAM', 'cell_column': 'G'},
    {'name': 'LABORATORY', 'cell_column': 'H'},
    {'name': 'END_DATE', 'cell_column': 'I'},
    {'name': 'NEXT_WORKS', 'cell_column': 'J'}
]

# Load Excel file
wb = openpyxl.load_workbook(excel_file)
ws = wb.active

# Find the last row with data
last_row = ws.max_row
while last_row > 0 and all(cell.value is None for cell in ws[last_row]):
    last_row -= 1

# Function to create a Word document
def create_word_doc(template_path, output_path, replacements):
    doc = Document(template_path)
    for para in doc.paragraphs:
        for run in para.runs:
            for key, value in replacements.items():
                if key in run.text:
                    run.text = run.text.replace(key, value)
    doc.save(output_path)

# Function to generate document from UI input
def generate_document():
    for row_number in range(4, 5): #range(4, last_row + 1):
        replacements = {}

        replacements[f"{main_data[0]['name']}"] = row0_entry.get()
        replacements[f"{main_data[1]['name']}"] = row1_entry.get()

        for obj in subobject_data:
            if obj['name'] == 'EXECUTION_DATE_MONTH':
                date_value = get_cell_value(ws, row_number, obj['cell_column']) or ""
                new_value = format_date(date_value)
                replacements[f"{obj['name']}"] = new_value
            else:    
                replacements[f"{obj['name']}"] = get_cell_value(ws, row_number, obj['cell_column']) or ""

            if obj['name'] == 'ACT_NUMBER':
                output_path = f"{get_cell_value(ws, row_number, obj['cell_column']).replace('/', '_')}.docx"
        
        try:
            create_word_doc(word_file, output_path, replacements)
            messagebox.showinfo("Успех", f"Документ сохранен как {output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

# UI Setup
root = tk.Tk()
root.title("Создание документов")

row0_value = get_cell_value(ws, main_data[0]['cell_row'], main_data[0]['cell_column']) or ""

tk.Label(root, text="Объект:").grid(row=0, column=0)
row0_entry = tk.Entry(root, width=150)
row0_entry.insert(0, row0_value)
row0_entry.grid(row=0, column=1)

row1_value = get_cell_value(ws, main_data[1]['cell_row'], main_data[1]['cell_column']) or ""

tk.Label(root, text="Субобъект:").grid(row=1, column=0)
row1_entry = tk.Entry(root, width=150)
row1_entry.insert(0, row1_value)
row1_entry.grid(row=1, column=1)

generate_button = tk.Button(root, text="Создать документ", command=generate_document)
generate_button.grid(row=2, columnspan=2)

root.mainloop()
