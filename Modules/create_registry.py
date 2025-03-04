from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import tkinter as tk
from tkinter import messagebox, ttk
import os
import re
from Modules.utils import get_cell_value, clear_ui, select_excel_file, format_date

REGISTER_TITLE_TEMPLATE = "Templates/Word_templates/Register_title_template.docx"
REGISTER_TABLE_TEMPLATE = "Templates/Word_templates/Register_table_template.docx"

MAIN_DATA = [
    {'name': 'OBJECTNAME', 'cell_row': '1', 'cell_column': 'B', 'description_cell_column': 'A'},
    {'name': 'CUSTOMER', 'cell_row': '14', 'cell_column': 'B', 'description_cell_column': 'A'},
    {'name': 'CONTRACTOR', 'cell_row': '15', 'cell_column': 'B', 'description_cell_column': 'A'},
    {'name': 'DESIGNORGANISATION', 'cell_row': '16', 'cell_column': 'B', 'description_cell_column': 'A'}
]

SUBOBJECT_DATA = [
    {'name': 'SUBOBJECTNAME', 'cell_column': 'B', 'cell_row': '1'},
    {'name': 'ACTNUMBER', 'cell_column': 'A'},
    {'name': 'EXECUTIONDATE', 'cell_column': 'B'},
    {'name': 'WORKNAMING', 'cell_column': 'C'},
    {'name': 'ALBUMNAME', 'cell_column': 'D'},
    {'name': 'PAGE', 'cell_column': 'E'},
    {'name': 'MATERIALS', 'cell_column': 'F'},
    {'name': 'EXECUTIVEDIAGRAM', 'cell_column': 'G'},
    {'name': 'LABORATORY', 'cell_column': 'H'},
    {'name': 'ENDDATE', 'cell_column': 'I'},
    {'name': 'NEXTWORKS', 'cell_column': 'J'}
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
        for i, value_data in enumerate(MAIN_DATA):
            label_text = get_cell_value(main_sheet, value_data['cell_row'], value_data['description_cell_column'])
            value = get_cell_value(main_sheet, value_data['cell_row'], value_data['cell_column'])
            
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
            if sheet_name not in {'Main data', 'Contents'} and not re.match(r'^[!_]', sheet_name):
                var = tk.BooleanVar()
                tk.Checkbutton(scrollable_frame, text=sheet_name, variable=var).grid(row=i, column=0, padx=5, pady=2, sticky="w")
                self.sheet_vars[sheet_name] = var

        canvas.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")

        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        self.all_var = tk.BooleanVar()
        tk.Checkbutton(frame, text="Выбрать все", variable=self.all_var, command=self.toggle_all_sheets).grid(row=0, column=0, padx=5, pady=5, sticky="w")

        generate_button = tk.Button(self.root, text="Создать реестр", command=self.generate_document)
        generate_button.grid(row=12, column=0, columnspan=2, pady=10)

    def save_to_excel(self):
        if self.wb is None:
            messagebox.showwarning("Предупреждение", "Сначала загрузите файл Excel!")
            return

        try:
            main_sheet = self.wb['Main data']
            for value_data in MAIN_DATA:
                cell = f"{value_data['cell_column']}{value_data['cell_row']}"
                main_sheet[cell].value = self.entries[value_data['name']].get()
            
            self.wb.save(self.excel_file_path)
            messagebox.showinfo("Успех", f"Данные сохранены в {self.excel_file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {str(e)}")

    def generate_document(self):
        if self.wb is None:
            messagebox.showwarning("Предупреждение", "Сначала загрузите файл Excel!")
            return

        selected_sheets = [sheet_name for sheet_name, var in self.sheet_vars.items() if var.get()]
        if not selected_sheets:
            messagebox.showwarning("Предупреждение", "Выберите хотя бы один лист!")
            return

        self.save_to_excel()

        created_files = []
        errors = []

        base_dir = os.path.join(os.getcwd(), "Documents")
        os.makedirs(base_dir, exist_ok=True)

        excel_dir = os.path.join(base_dir, self.excel_file_name)
        os.makedirs(excel_dir, exist_ok=True)

        for sheet_name in selected_sheets:
            ws = self.wb[sheet_name]
            last_row = ws.max_row
            while last_row > 0 and all(cell.value is None for cell in ws[last_row]):
                last_row -= 1

            sheet_dir = os.path.join(excel_dir, sheet_name)
            os.makedirs(sheet_dir, exist_ok=True)

            output_path = None  # Инициализируем переменную перед try
            try:
                output_path, doc = self.generate_document_title(ws, last_row, sheet_dir)
                doc = self.generate_document_table(doc, ws, last_row)
                self.generate_document_content()  # Пока пустая, можно убрать если не используется
                
                doc.save(output_path)
                created_files.append(output_path)
            except Exception as e:
                error_path = output_path if output_path else f"для листа {sheet_name}"
                errors.append(f"Ошибка при создании {error_path}: {str(e)}")

        if created_files:
            messagebox.showinfo("Успех", f"Созданы файлы:\n" + "\n".join(created_files))
        if errors:
            messagebox.showerror("Ошибки", "\n".join(errors))

    def generate_document_title(self, ws, last_row, sheet_dir):
        replacements = {}
        for value_data in MAIN_DATA:
            replacements[value_data['name']] = self.entries[value_data['name']].get()

        for obj in SUBOBJECT_DATA:
            if obj['name'] == 'SUBOBJECTNAME':
                replacements[obj['name']] = get_cell_value(ws, obj['cell_row'], obj['cell_column'])
            elif obj['name'] == 'EXECUTIONDATE':
                replacements[obj['name']] = get_cell_value(ws, 3, obj['cell_column'])
            elif obj['name'] == 'ENDDATE':
                replacements[obj['name']] = get_cell_value(ws, last_row, obj['cell_column'])

        output_path = os.path.join(sheet_dir, f"Реестр {replacements['SUBOBJECTNAME']}.docx")
        doc = self.create_word_doc(REGISTER_TITLE_TEMPLATE, replacements)
        return output_path, doc        

    # Пример интеграции в ваш код
    # def generate_document_table(self, doc, ws, last_row):
    #     # Добавляем разрыв страницы перед таблицей
    #     doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        
    #     table_doc = Document(REGISTER_TABLE_TEMPLATE)
    #     for table in table_doc.tables:
    #         # Определяем количество необходимых строк и их контент
    #         row_idx = 1
    #         page = 0
    #         table_data = {}
            
    #         # Заполняем хеш-таблицу данными из ws
    #         for i in range(3, last_row + 1):  # Начинаем с 3, заканчиваем last_row
    #             # Инициализируем текущую строку в table_data
    #             table_data[row_idx] = {}
    #             table_data[row_idx]['idx'] = str(row_idx)
    #             act_number = get_cell_value(ws, i, SUBOBJECT_DATA[1]['cell_column']) or ""
    #             work_naming = get_cell_value(ws, i, SUBOBJECT_DATA[3]['cell_column']) or ""
    #             exec_date = get_cell_value(ws, i, SUBOBJECT_DATA[2]['cell_column']) or ""
    #             table_data[row_idx]['content'] = f"Акт скрытых работ № {act_number} {work_naming} {format_date(exec_date) if exec_date else ''}"
    #             page += 2
    #             table_data[row_idx]['page'] = page

    #             # Проверяем material
    #             material = get_cell_value(ws, i, SUBOBJECT_DATA[6]['cell_column']) or ""
    #             if material != '':
    #                 row_idx += 1
    #                 table_data[row_idx] = {}
    #                 table_data[row_idx]['idx'] = str(row_idx)
    #                 table_data[row_idx]['content'] = material
    #                 page += 2
    #                 table_data[row_idx]['page'] = page

    #             # Проверяем schema
    #             schema = get_cell_value(ws, i, SUBOBJECT_DATA[7]['cell_column']) or ""
    #             if schema != '':
    #                 row_idx += 1
    #                 table_data[row_idx] = {}
    #                 table_data[row_idx]['idx'] = str(row_idx)
    #                 table_data[row_idx]['content'] = schema  # Исправлено: schema вместо material
    #                 page += 2
    #                 table_data[row_idx]['page'] = page

    #             # Проверяем laboratory
    #             laboratory = get_cell_value(ws, i, SUBOBJECT_DATA[8]['cell_column']) or ""
    #             if laboratory != '':
    #                 row_idx += 1
    #                 table_data[row_idx] = {}
    #                 table_data[row_idx]['idx'] = str(row_idx)
    #                 table_data[row_idx]['content'] = laboratory  # Исправлено: laboratory вместо material
    #                 page += 2
    #                 table_data[row_idx]['page'] = page

    #             row_idx += 1

    #         data_rows_needed = row_idx - 1  # Учитываем, что row_idx увеличивается лишний раз
    #         print(f"Data rows needed: {data_rows_needed}")

    #         template_rows = len(table.rows)
            
    #         # Создаем новую таблицу
    #         new_table = doc.add_table(rows=max(template_rows, data_rows_needed + 1), cols=len(table.columns))
            
    #         # Копируем свойства таблицы из шаблона
    #         new_table.autofit = table.autofit
            
    #         # Копируем ширину колонок
    #         for col_idx, column in enumerate(table.columns):
    #             if col_idx < len(new_table.columns):
    #                 new_table.columns[col_idx].width = column.width
            
    #         # Копируем содержимое и форматирование ячеек из шаблона
    #         for i, row in enumerate(table.rows):
    #             new_row = new_table.rows[i]
    #             if row._tr.trPr.find(qn('w:trHeight')) is not None:
    #                 new_row._tr.trPr.append(row._tr.trPr.find(qn('w:trHeight')))
    #             for j, cell in enumerate(row.cells):
    #                 new_cell = new_row.cells[j]
    #                 new_cell.text = cell.text
    #                 new_cell.width = cell.width
    #                 set_cell_borders(new_cell)  # Добавляем границы
    #                 for src_para, dst_para in zip(cell.paragraphs, new_cell.paragraphs):
    #                     dst_para.style = src_para.style
    #                     dst_para.paragraph_format.space_before = 0
    #                     dst_para.paragraph_format.space_after = 0
    #                     dst_para.paragraph_format.line_spacing = 1.0
    #                     for src_run, dst_run in zip(src_para.runs, dst_para.runs):
    #                         dst_run.font.name = src_run.font.name
    #                         dst_run.font.size = src_run.font.size
    #                         dst_run.bold = src_run.bold
    #                         dst_run.italic = src_run.italic
    #                         dst_run.underline = src_run.underline

    #         # Заполняем таблицу данными из table_data
    #         for row_idx in range(1, data_rows_needed):
    #             if row_idx >= template_rows:
    #                 new_row = new_table.add_row()
    #                 template_row = table.rows[1]
    #                 if template_row._tr.trPr.find(qn('w:trHeight')) is not None:
    #                     new_row._tr.trPr.append(template_row._tr.trPr.find(qn('w:trHeight')))
    #             cells = new_table.rows[row_idx].cells
    #             cells[0].text = table_data[row_idx]['idx']  # Номер строки
    #             cells[1].text = table_data[row_idx]['content']  # Контент
    #             cells[2].text = str(table_data[row_idx]['page'])  # Страница
    #             for cell in cells:
    #                 set_cell_borders(cell)  # Добавляем границы
    #                 for para in cell.paragraphs:
    #                     para.paragraph_format.space_before = 0
    #                     para.paragraph_format.space_after = 0
    #                     para.paragraph_format.line_spacing = 1.0
            
    #         # Устанавливаем отступы для таблицы
    #         table_paragraph = new_table._element.getparent()
    #         if table_paragraph.tag.endswith('p'):
    #             table_paragraph.paragraph_format.left_indent = 250000  # 1 см слева
    #             table_paragraph.paragraph_format.right_indent = 250000  # 1 см справа
            
    #         new_table.autofit = False
        
    #     return doc

    def generate_document_table(self, doc, ws, last_row):
        # Добавляем разрыв страницы перед таблицей
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        
        table_doc = Document(REGISTER_TABLE_TEMPLATE)
        for table in table_doc.tables:
            # Заполняем данные таблицы
            table_data, data_rows_needed = fill_table_data(ws, last_row)            
            template_rows = len(table.rows)
            
            # Создаем новую таблицу
            new_table = doc.add_table(rows=max(template_rows, data_rows_needed), cols=len(table.columns))
            
            # Копируем форматирование из шаблона
            copy_template_formatting(new_table, table)
            
            # Заполняем таблицу данными
            fill_table_with_data(new_table, table_data, template_rows, table)
            
            # Устанавливаем отступы для таблицы
            table_paragraph = new_table._element.getparent()
            if table_paragraph.tag.endswith('p'):
                table_paragraph.paragraph_format.left_indent = 250000  # 1 см слева
                table_paragraph.paragraph_format.right_indent = 250000  # 1 см справа
            
            new_table.autofit = False
        
        return doc

    def generate_document_content(self):
        pass

    def create_word_doc(self, template_path, replacements):
        doc = Document(template_path)
        for para in doc.paragraphs:
            for run in para.runs:
                for key, value in replacements.items():
                    if key == run.text:
                        run.text = run.text.replace(key, value)
        return doc

    def toggle_all_sheets(self):
        select_all = self.all_var.get()
        for var in self.sheet_vars.values():
            var.set(select_all)

def set_cell_borders(cell):
    """Устанавливает видимые границы для ячейки."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), '000000')
        borders.append(border)
    tcPr.append(borders)

def add_table_row_data(table_data, row_idx, content, page):
    table_data[row_idx] = {
        'idx': str(row_idx),
        'content': content,
        'page': page
    }
    return row_idx + 1

def process_split_text(table_data, row_idx, text, page):
    if ';' in text:
        parts = [part.strip() for part in text.split(';') if part.strip()]
        for part in parts:
            page += 2
            row_idx = add_table_row_data(table_data, row_idx, part, page)
    elif text != '':
        page += 2
        row_idx = add_table_row_data(table_data, row_idx, text, page)
    return row_idx, page

def fill_table_data(ws, last_row):
    """Заполняет хеш-таблицу данными из ws."""
    row_idx = 1
    page = 0
    table_data = {}
    
    for i in range(3, last_row + 1):
        # Основная строка с актом
        act_number = get_cell_value(ws, i, SUBOBJECT_DATA[1]['cell_column']) or ""
        work_naming = get_cell_value(ws, i, SUBOBJECT_DATA[3]['cell_column']) or ""
        exec_date = get_cell_value(ws, i, SUBOBJECT_DATA[2]['cell_column']) or ""
        content = f"Акт скрытых работ № {act_number} {work_naming} {format_date(exec_date) if exec_date else ''}"
        page += 2
        row_idx = add_table_row_data(table_data, row_idx, content, page)

        # Проверяем material
        material = get_cell_value(ws, i, SUBOBJECT_DATA[6]['cell_column']) or ""
        row_idx, page = process_split_text(table_data, row_idx, material, page)

        # Проверяем schema
        schema = get_cell_value(ws, i, SUBOBJECT_DATA[7]['cell_column']) or ""
        row_idx, page = process_split_text(table_data, row_idx, schema, page)

        # Проверяем laboratory
        laboratory = get_cell_value(ws, i, SUBOBJECT_DATA[8]['cell_column']) or ""
        row_idx, page = process_split_text(table_data, row_idx, laboratory, page)

    return table_data, row_idx - 1

def copy_template_formatting(new_table, template_table):
    """Копирует форматирование из шаблона в новую таблицу."""
    new_table.autofit = template_table.autofit
    for col_idx, column in enumerate(template_table.columns):
        if col_idx < len(new_table.columns):
            new_table.columns[col_idx].width = column.width
    
    for i, row in enumerate(template_table.rows):
        new_row = new_table.rows[i]
        if row._tr.trPr.find(qn('w:trHeight')) is not None:
            new_row._tr.trPr.append(row._tr.trPr.find(qn('w:trHeight')))
        for j, cell in enumerate(row.cells):
            new_cell = new_row.cells[j]
            new_cell.text = cell.text
            new_cell.width = cell.width
            set_cell_borders(new_cell)
            for src_para, dst_para in zip(cell.paragraphs, new_cell.paragraphs):
                dst_para.style = src_para.style
                dst_para.paragraph_format.space_before = 0
                dst_para.paragraph_format.space_after = 0
                dst_para.paragraph_format.line_spacing = 1.0
                for src_run, dst_run in zip(src_para.runs, dst_para.runs):
                    dst_run.font.name = src_run.font.name
                    dst_run.font.size = src_run.font.size
                    dst_run.bold = src_run.bold
                    dst_run.italic = src_run.italic
                    dst_run.underline = src_run.underline

def fill_table_with_data(new_table, table_data, template_rows, template_table):
    """Заполняет таблицу данными из хеш-таблицы."""
    for row_idx in range(1, len(table_data) + 1):
        if row_idx >= template_rows:
            new_row = new_table.add_row()
            template_row = template_table.rows[1]
            if template_row._tr.trPr.find(qn('w:trHeight')) is not None:
                new_row._tr.trPr.append(template_row._tr.trPr.find(qn('w:trHeight')))
        cells = new_table.rows[row_idx].cells
        cells[0].text = table_data[row_idx]['idx']
        cells[1].text = table_data[row_idx]['content']
        cells[2].text = str(table_data[row_idx]['page'])
        for cell in cells:
            set_cell_borders(cell)
            for para in cell.paragraphs:
                para.paragraph_format.space_before = 0
                para.paragraph_format.space_after = 0
                para.paragraph_format.line_spacing = 1.0