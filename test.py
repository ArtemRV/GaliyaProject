# from pdf2image import convert_from_path
# import pytesseract
# import re
# import os
# from pathlib import Path

# input_dir = Path(r"C:\Users\Artsiom_Rachok\Downloads\Telegram Desktop\Новая папка\Новая папка")
# output_dir = input_dir / "новые"
# output_dir.mkdir(exist_ok=True)

# pattern = re.compile(r"Протокол\s*испытаний\s*№\s*[\d/-]+", re.IGNORECASE)

# def sanitize_filename(name: str) -> str:
#     # Заменяем слеши
#     name = name.replace("/", "_").replace("\\", "_")

#     # Запрещённые символы Windows:  <>:"/\|?*
#     forbidden = r'<>:"/\|?*'
#     for ch in forbidden:
#         name = name.replace(ch, "_")

#     # Убираем повторяющиеся пробелы
#     name = re.sub(r"\s+", " ", name).strip()

#     return name


# for pdf_path in input_dir.glob("*.pdf"):
#     pages = convert_from_path(pdf_path, dpi=200)
#     text = ""

#     for page in pages:
#         text += pytesseract.image_to_string(page, lang="rus")

#     match = pattern.search(text)
#     if match:
#         clean = sanitize_filename(match.group(0))
#         new_name = f"{clean}.pdf"
#     else:
#         new_name = pdf_path.name

#     os.rename(pdf_path, output_dir / new_name)
#     print(f"{pdf_path.name} → {new_name}")


from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Pattern, Dict

import re
import os
from pdf2image import convert_from_path
import pytesseract


# ---------------------------
# 🔧 Утилита очистки имени файла
# ---------------------------
def sanitize_filename(name: str) -> str:
    name = name.replace("/", "_").replace("\\", "_")
    forbidden = r'<>:"/\|?*'
    for ch in forbidden:
        name = name.replace(ch, "_")
    name = re.sub(r"\s+", " ", name).strip()
    return name


# ---------------------------
# 🔧 Конфигурация
# ---------------------------
@dataclass
class RenameRule:
    """Описывает правило поиска и формат результата."""
    pattern: Pattern
    template: str  # например: "{match}.pdf"


@dataclass
class PDFRenamerConfig:
    input_dirs: List[Path]
    output_dir: Path
    rules: List[RenameRule] = field(default_factory=list)
    dpi: int = 200
    lang: str = "rus"


# ---------------------------
# 🔧 Основной класс
# ---------------------------
class PDFRenamer:

    def __init__(self, config: PDFRenamerConfig):
        self.config = config
        self.config.output_dir.mkdir(parents=True, exist_ok=True)

    def extract_text(self, pdf_path: Path) -> str:
        """Конвертирует PDF в изображения, затем в текст."""
        pages = convert_from_path(pdf_path, dpi=self.config.dpi)
        text = ""

        for page in pages:
            text += pytesseract.image_to_string(page, lang=self.config.lang)

        return text

    def apply_rules(self, text: str) -> str | None:
        """Возвращает имя файла по первому совпавшему правилу."""
        for rule in self.config.rules:
            match = rule.pattern.search(text)
            if match:
                filename = rule.template.format(match=sanitize_filename(match.group(0)))
                return filename
        return None

    def process_file(self, pdf_path: Path):
        """Обрабатывает один PDF."""
        print(f"📄 Обработка: {pdf_path.name}")

        text = self.extract_text(pdf_path)
        new_name = self.apply_rules(text)

        if not new_name:
            print("⚠ Совпадений нет — сохраняем исходное имя")
            new_name = pdf_path.name

        target = self.config.output_dir / new_name
        os.rename(pdf_path, target)

        print(f"✔ Переименован → {target.name}\n")

    def run(self):
        """Запуск обработки всех файлов из всех директорий."""
        for directory in self.config.input_dirs:
            print(f"📁 Чтение папки: {directory}")

            for pdf_path in directory.glob("*.pdf"):
                self.process_file(pdf_path)


# ---------------------------
# 🔧 Пример использования
# ---------------------------
if __name__ == "__main__":

    config = PDFRenamerConfig(
        input_dirs=[
            Path(r"C:\Users\Artsiom_Rachok\Downloads\Telegram Desktop\Новая папка\Новая папка")
        ],
        output_dir=Path(r"C:\Users\Artsiom_Rachok\Downloads\Telegram Desktop\Новая папка\Новая папка\новые"),

        rules=[
            RenameRule(
                pattern=re.compile(r"Протокол\s*испытаний\s*№\s*[\d/-]+", re.IGNORECASE),
                template="{match}.pdf"
            ),

            # можно легко добавлять новые паттерны:
            RenameRule(
                pattern=re.compile(r"Договор\s*№\s*[\d-]+", re.IGNORECASE),
                template="DOGOVOR_{match}.pdf"
            )
        ],

        dpi=200,
        lang="rus"
    )

    renamer = PDFRenamer(config)
    renamer.run()
