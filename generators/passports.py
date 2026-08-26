import os
import pandas as pd
from copy import deepcopy
from docxtpl import DocxTemplate
from docx import Document


def generate_passports(df, template_path, output_path, start_row, end_row):
    try:

        rows = df.iloc[start_row - 1:end_row].to_dict(orient="records")

        base_template = DocxTemplate(template_path)
        base_template.render({})
        base_template.save("~temp_base.docx")

        final_doc = Document("~temp_base.docx")

        # Очищаем тело финального документа
        for element in final_doc.element.body[:]:
            final_doc.element.body.remove(element)

        generated_count = 0

        for i, context in enumerate(rows, start=1):
            try:
                template = DocxTemplate(template_path)
                template.render(context)
                template.save("~temp_passport.docx")

                temp_doc = Document("~temp_passport.docx")

                for elem in temp_doc.element.body:
                    # Исключаем пустые абзацы и служебные элементы (например, w:sectPr)
                    if elem.tag.endswith("sectPr"):
                        continue
                    if elem.tag.endswith("p"):  # Параграф
                        texts = [node.text for node in elem.iter() if node.text and node.text.strip()]
                        if not texts:
                            continue  # пропустить пустые параграфы

                    final_doc.element.body.append(deepcopy(elem))

                if i < len(rows):
                    final_doc.add_page_break()

                generated_count += 1

            except Exception as e:
                print(f"Ошибка при создании паспорта {i}: {str(e)}")

        final_doc.save(output_path)
        os.remove("~temp_base.docx")
        os.remove("~temp_passport.docx")

        return True, f"Сгенерировано {generated_count} паспортов в одном файле: {output_path}"

        # Сохраняем финальный файл
        final_doc.save(output_path)
        os.remove("~temp_base.docx")
        os.remove("~temp_passport.docx")
        return True, f"Сгенерировано {generated_count} паспортов в одном файле: {output_path}"

    except Exception as e:
        return False, f"Ошибка при создании паспортов: {str(e)}"