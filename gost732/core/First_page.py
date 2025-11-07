from docxtpl import DocxTemplate
from pathlib import Path
from datetime import datetime
import sys

current_date = datetime.now()

# --- НАСТРОЙКА: Определение корня проекта ---
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.append(str(PROJECT_ROOT))

# Убедимся, что папка reports существует
REPORTS_DIR = PROJECT_ROOT / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

# 1. Определяем пути к файлам
TEMPLATE_PATH = PROJECT_ROOT / "template" / "Main_Template.docx"
OUTPUT_PATH = REPORTS_DIR / "Lab_1_Ivanov.docx"

doc = DocxTemplate(TEMPLATE_PATH)

context = {
    'номер_кафедры': 12,
    'должность_преподавателя': 'ассистент',
    'ФИО_преподавателя': 'Сидоров С.С.',
    'номер_работы': 1,
    'название_работы': 'Исследование чего-то там',
    'ФИО_студента': 'Иванов И.И.',
    'группа': '4300',
    'ПР_ЛБ': 'ЛАБОРАТОРНОЙ',
    'название_дисциплины': 'Программирование',
    'выполнил_выполнила': 'ВЫПОЛНИЛ',
    'студент_студентка': 'Студент',
    'дата': datetime.strftime(current_date, "%y.%m.%d"),
    'год': str(current_date.year),

    'основной текст': ''
}

doc.render(context)
doc.save(OUTPUT_PATH)

print(f"Документ сохранен в: {OUTPUT_PATH}")