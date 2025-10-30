from pathlib import Path
import yaml
from docx import Document
from docx.shared import Pt, Cm, Inches # Для работы с размерами (точки, см, дюймы)
from docx.enum.text import WD_ALIGN_PARAGRAPH # Для работы с выравниванием (CENTER, LEFT и т.д.)
import re # Для работы с регулярными выражениями (для рисунков/таблиц)

# Определяем корневую папку проекта
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def load_gost_rules(rules_file: str = "rules.yaml"):
    """Загружает правила ГОСТа из YAML файла."""
    rules_path = PROJECT_ROOT / "src" / rules_file

    if not rules_path.exists():
        print(f"Ошибка: Файл правил не найден по пути {rules_path}")
        return None

    with open(rules_path, 'r', encoding='utf-8') as f:
        rules = yaml.safe_load(f)

    print(f"Правила успешно загружены: {list(rules.keys())}")
    return rules


def check_main_text_format(paragraph, rules, index):
    """Проверяет форматирование абзаца основного текста по ГОСТу."""

    reqs = rules['formatting_requirements']
    required_indent = Cm(reqs['first_line_indent_cm'])
    required_font_size = Pt(reqs['font_size_pt'])

    # Считаем, что все, что не заголовок - это основной текст (для простоты MVP)
    if paragraph.style.name == rules['main_text_style']:

        # --- ИСПРАВЛЕННАЯ ПРОВЕРКА АБЗАЦНОГО ОТСТУПА ---
        current_indent = paragraph.paragraph_format.first_line_indent

        # 1. Если отступ None, мы считаем, что он равен 0 (для Word это логично)
        if current_indent is None:
            current_indent_cm = 0.0
        else:
            current_indent_cm = current_indent.cm

        # 2. Сравниваем, используя значение в см
        if abs(current_indent_cm - required_indent.cm) > 0.01:
            print(
                f"[{index}] ❌ Ошибка отступа: Ожидается {required_indent.cm:.2f} см (стиль '{paragraph.style.name}'), Найдено {current_indent_cm:.2f} см")

        # -------------------------------------------------------------------

        # Проверка размера шрифта (проверяем только первый Run в абзаце для MVP)
        if paragraph.runs and paragraph.runs[0].font.size != required_font_size:
            current_size = paragraph.runs[0].font.size.pt if paragraph.runs[0].font.size else "N/A"
            print(
                f"[{index}] ❌ Ошибка размера: Ожидается {required_font_size.pt} pt, Найдено {current_size} pt (стиль '{paragraph.style.name}')")

    # NOTE: Если абзац не основного стиля, мы пока его игнорируем,
    # но в будущем нам нужно будет проверять его на соответствие своему стилю
    # (например, Заголовок 1)


def validate_document(filepath: Path, rules: dict):
    """Основная логика валидации."""

    if not filepath.exists():
        print(f"⛔️ Ошибка: Файл для проверки не найден по пути {filepath}")
        return

    print(f"\n--- 🕵️‍♂️ Проверка документа: {filepath.name} ---")
    doc = Document(filepath)

    # Итерируемся по всем абзацам документа
    for i, p in enumerate(doc.paragraphs):
        # Проверка №1: Стиль основного текста
        check_main_text_format(p, rules, i + 1)  # i + 1 для нумерации с 1

        # Проверка №2 (TODO: рисунки, списки и т.д. будут добавлены здесь)

    print("--- Проверка завершена. ---")


if __name__ == "__main__":
    # 1. Загружаем правила
    gost_rules = load_gost_rules()

    if gost_rules:
        # 2. Определяем файл для проверки (возьмем наш сгенерированный)
        target_file = PROJECT_ROOT / "reports" / "Lab_1_Ivanov.docx"

        # 3. Запускаем валидацию
        validate_document(target_file, gost_rules)