from docx import Document


def open_document(template_path):
    """Открывает шаблон Word и возвращает документ."""
    return Document(template_path)


def get_first_table(doc):
    """Возвращает первую таблицу документа."""
    return doc.tables[0]


def save_document(doc, output_path):
    """Сохраняет документ."""
    doc.save(output_path)