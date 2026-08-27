from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from docxtpl import DocxTemplate

from config import WorkflowConfig
from workflow_adapters import get_passport_context_adapter


def _clear_document_body(document):
    """Remove template content while preserving its section properties."""

    body = document.element.body
    section_properties = next(
        (
            deepcopy(element)
            for element in body
            if element.tag.endswith("sectPr")
        ),
        None,
    )

    for element in list(body):
        body.remove(element)

    return section_properties


def _append_passport(final_document, passport_document):
    """Append one rendered passport without its section properties."""

    for element in passport_document.element.body:
        if element.tag.endswith("sectPr"):
            continue

        if element.tag.endswith("p"):
            texts = [
                node.text
                for node in element.iter()
                if node.text and node.text.strip()
            ]
            if not texts:
                continue

        final_document.element.body.append(deepcopy(element))


def generate_passports(
    df,
    workflow: WorkflowConfig,
    output_path,
    start_row,
    end_row,
):
    """Render selected records with the configured workflow adapter."""

    try:
        rows = df.iloc[start_row - 1:end_row].to_dict(orient="records")

        if not rows:
            raise ValueError(
                "В выбранном диапазоне нет записей для создания паспортов."
            )

        context_adapter = get_passport_context_adapter(
            workflow.passport_context_adapter
        )

        with TemporaryDirectory(prefix="inventory-passports-") as temp_dir:
            temp_path = Path(temp_dir)
            base_path = temp_path / "base.docx"

            base_template = DocxTemplate(workflow.passport_template)
            base_template.render({})
            base_template.save(base_path)

            final_document = Document(base_path)
            section_properties = _clear_document_body(final_document)

            for index, record in enumerate(rows, start=1):
                try:
                    context = context_adapter(record)
                    rendered_path = temp_path / f"passport-{index}.docx"

                    template = DocxTemplate(workflow.passport_template)
                    template.render(context)
                    template.save(rendered_path)

                    _append_passport(
                        final_document,
                        Document(rendered_path),
                    )
                except Exception as exc:
                    raise RuntimeError(
                        f"ошибка при создании паспорта {index}: {exc}"
                    ) from exc

                if index < len(rows):
                    final_document.add_page_break()

            if section_properties is not None:
                final_document.element.body.append(section_properties)

            final_document.save(output_path)

        return (
            True,
            f"Сгенерировано {len(rows)} паспортов в одном файле: "
            f"{output_path}",
        )

    except Exception as exc:
        return False, f"Ошибка при создании паспортов: {exc}"
