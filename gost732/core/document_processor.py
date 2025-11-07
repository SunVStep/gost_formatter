from docx import Document
from pathlib import Path

project_path = Path(__file__).resolve().parent.parent
reports_dir = project_path / 'reports'
templates_dir = project_path / 'template'

dirty_path = reports_dir / 'test.docx'
template_path = templates_dir / 'Main_Template.docx'
output_path = reports_dir / 'cleaned_document.docx'  # Путь для сохранения результата

# Проверка существования файлов
if not dirty_path.exists():
    raise FileNotFoundError(f"Исходный файл не найден: {dirty_path}")
if not template_path.exists():
    raise FileNotFoundError(f"Файл шаблона не найден: {template_path}")

dirty_doc = Document(dirty_path)
final_document = Document(template_path)


def insert_text(paragraph, text, bold=False, italic=False, underline=False):
    """Добавляет текст в параграф с сохранением форматирования"""
    run = paragraph.add_run(text)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if underline is not None:
        run.underline = underline
    return run


def process_document(source_doc: Document, target_doc: Document):
    """Обрабатывает исходный документ и добавляет контент в целевой"""
    for paragraph in source_doc.paragraphs:
        # Создаём новый параграф в целевом документе
        new_paragraph = target_doc.add_paragraph()

        # Применяем стиль ГОСТ_Текст
        if paragraph.style:
            try:
                new_paragraph.style = "ГОСТ_Текст"
            except KeyError:
                pass

        # Обрабатываем каждый ран в параграфе
        for run in paragraph.runs:
            insert_text(
                new_paragraph,
                run.text,
                bold=run.bold,
                italic=run.italic,
                underline=run.underline
            )


# Основная обработка
process_document(dirty_doc, final_document)

# Сохраняем результат
final_document.save(output_path)
print(f"Документ успешно обработан и сохранён в: {output_path}")