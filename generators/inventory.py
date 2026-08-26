from copy import copy

from openpyxl import load_workbook

from generators.inventory_fields import REPORT_FIELDS


# ==========================================================
# Константы
# ==========================================================

FIRST_DATA_ROW = 17
TEMPLATE_ROW = 17


# ==========================================================
# Главная функция
# ==========================================================

def generate_inventory_report(
    df,
    template_path,
    output_path,
    start_row,
    end_row,
):
    """
    Генерирует инвентаризационный отчет.
    """

    try:

        workbook = open_template(template_path)
        sheet = workbook.active

        rows = select_rows(
            df,
            start_row,
            end_row,
        )

        fill_sheet(
            sheet,
            rows,
        )

        finalize_sheet(sheet)

        save_report(
            workbook,
            output_path,
        )

        return (
            True,
            f"Инвентаризационный отчет сохранен:\n{output_path}",
        )

    except Exception as e:

        return (
            False,
            f"Ошибка генерации отчета:\n{e}",
        )


# ==========================================================
# Работа с книгой
# ==========================================================

def open_template(template_path):
    """
    Открывает шаблон Excel.
    """

    return load_workbook(template_path)


def save_report(workbook, output_path):
    """
    Сохраняет отчет.
    """

    workbook.save(output_path)


# ==========================================================
# Работа с данными
# ==========================================================

def select_rows(
    df,
    start_row,
    end_row,
):
    """
    Возвращает выбранные пользователем строки.
    """

    return df.iloc[start_row - 1:end_row]


# ==========================================================
# Работа со строками шаблона
# ==========================================================

def insert_data_row(
    sheet,
    row_number,
):
    """
    Вставляет новую строку
    и полностью копирует оформление строки шаблона.
    """

    sheet.insert_rows(row_number)

    copy_row_style(
        sheet,
        TEMPLATE_ROW,
        row_number,
    )

    copy_row_dimensions(
        sheet,
        TEMPLATE_ROW,
        row_number,
    )

    copy_row_merges(
        sheet,
        TEMPLATE_ROW,
        row_number,
    )


def copy_row_style(
    sheet,
    source_row,
    target_row,
):
    """
    Копирует стили строки.
    """

    for cell in sheet[source_row]:

        target = sheet[
            f"{cell.column_letter}{target_row}"
        ]

        if cell.has_style:
            target._style = copy(cell._style)

        target.font = copy(cell.font)
        target.fill = copy(cell.fill)
        target.border = copy(cell.border)
        target.alignment = copy(cell.alignment)
        target.protection = copy(cell.protection)
        target.number_format = copy(cell.number_format)


def copy_row_dimensions(
    sheet,
    source_row,
    target_row,
):
    """
    Копирует высоту строки.
    """

    sheet.row_dimensions[target_row].height = (
        sheet.row_dimensions[source_row].height
    )


def copy_row_merges(
    sheet,
    source_row,
    target_row,
):
    """
    Копирует объединенные ячейки строки.
    """

    ranges = list(sheet.merged_cells.ranges)

    for merged in ranges:

        if (
            merged.min_row == source_row
            and merged.max_row == source_row
        ):

            sheet.merge_cells(

                start_row=target_row,
                start_column=merged.min_col,

                end_row=target_row,
                end_column=merged.max_col,
            )

# ==========================================================
# Заполнение листа
# ==========================================================

def fill_sheet(sheet, rows):
    """
    Заполняет отчет данными.
    """

    current_row = FIRST_DATA_ROW
    current_number = 1

    for data in rows.itertuples(index=False):

        insert_data_row(
            sheet,
            current_row,
        )

        fill_row(
            sheet=sheet,
            sheet_row=current_row,
            number=current_number,
            data=data,
        )

        current_row += 1
        current_number += 1


def fill_row(
    sheet,
    sheet_row,
    number,
    data,
):
    """
    Заполняет одну строку отчета.
    """

    write_cell(
        sheet,
        f"A{sheet_row}",
        number,
    )

    for column, builder in REPORT_FIELDS:

        value = builder(data)

        write_cell(
            sheet,
            f"{column}{sheet_row}",
            value,
        )


def write_cell(
    sheet,
    cell,
    value,
):
    """
    Записывает значение в ячейку.
    """

    if value is None:
        value = ""

    sheet[cell] = value


# ==========================================================
# Финальная обработка
# ==========================================================

def finalize_sheet(sheet):
    """
    Здесь можно будет разместить финальные действия
    перед сохранением книги.
    """

    pass