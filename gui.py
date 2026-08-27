import FreeSimpleGUI as sg

from config import get_departments, get_workflow
from data_loader import InputSchemaError, load_excel
from generators.inventory import generate_inventory_report
from generators.passports import generate_passports
from record_selection import (
    RecordSelectionError,
    parse_row_range,
    select_record_range,
)
from workflow_context import InventoryContext, calculate_staff_count


def load_source_data(path):
    """Load an Excel source and return a user-facing error instead of raising."""

    try:
        return load_excel(path), None
    except FileNotFoundError:
        return None, "Excel-файл не найден."
    except PermissionError:
        return None, "Нет доступа к Excel-файлу."
    except InputSchemaError as exc:
        detail = str(exc).removeprefix(
            "Некорректная схема входного Excel: "
        )
        return (
            None,
            "Файл не соответствует ожидаемой структуре: " + detail,
        )
    except ValueError:
        return (
            None,
            "Не удалось прочитать Excel-файл. "
            "Проверьте его формат и содержимое.",
        )
    except OSError:
        return None, "Не удалось открыть Excel-файл."
    except Exception:
        return None, "Не удалось загрузить Excel-файл."


def run():
    sg.theme("LightGrey1")

    layout = [
        [
            sg.Text("Excel-файл:"),
            sg.Input(key="-EXCEL-"),
            sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))
        ],

        [
            sg.Frame(
                "Инвентаризационный отчёт",
                [
                    [
                        sg.Text("Наименование объекта:"),
                        sg.Input(key="-OBJECT_NAME-")
                    ],
                    [
                        sg.Text("Адрес:"),
                        sg.Input(key="-OBJECT_ADDRESS-")
                    ],
                    [
                        sg.Text("Куда сохранить:"),
                        sg.Input(key="-INV_OUTPUT-"),
                        sg.FileSaveAs(file_types=(("Excel Files", "*.xlsx"),))
                    ]
                ]
            )
        ],

        [
            sg.Frame(
                "Паспорта оборудования",
                [
                    [
                        sg.Text("Тип подразделения:"),
                        sg.Combo(
                        values=get_departments(),
                        default_value=get_departments()[0],
                        readonly=True,
                        key="-DEPARTMENT-"
                    )
                    ],

                    [
                        sg.Text("Куда сохранить:"),
                        sg.Input(key="-PASSP_OUTPUT-"),
                        sg.FileSaveAs(file_types=(("Word Files", "*.docx"),))
                    ]
                ]
            )
        ],

        [
            sg.Text("Строки (от/до):"),
            sg.Input("1", size=5, key="-START-"),
            sg.Input("5", size=5, key="-END-")
        ],

        [
            sg.Checkbox(
                "Генерировать инвентаризационный отчёт",
                default=True,
                key="-GEN_INV-"
            ),

            sg.Checkbox(
                "Генерировать паспорта",
                default=True,
                key="-GEN_PASSP-"
            )
        ],

        [
            sg.Button("Сгенерировать"),
            sg.Button("Выход")
        ],

        [
            sg.Output(size=(80, 10), key="-LOG-")
        ]
    ]

    window = sg.Window("Генератор отчётов и паспортов", layout)

    while True:
        event, values = window.read()

        if event in (None, "Выход"):
            break

        if event != "Сгенерировать":
            continue

        # ---------- Проверки ----------

        if not values["-EXCEL-"]:
            sg.popup_error("Укажите Excel-файл!")
            continue

        try:
            start_row, end_row = parse_row_range(
                values["-START-"],
                values["-END-"],
            )
        except RecordSelectionError as exc:
            sg.popup_error(str(exc))
            continue

        gen_inv = values["-GEN_INV-"]
        gen_passp = values["-GEN_PASSP-"]
        generation_succeeded = True

        if not gen_inv and not gen_passp:
            sg.popup_error("Выберите хотя бы один тип документа!")
            continue

        if gen_inv and not values["-INV_OUTPUT-"]:
            sg.popup_error("Укажите путь сохранения инвентаризационного отчёта!")
            continue

        if gen_inv and not values["-OBJECT_NAME-"]:
            sg.popup_error("Укажите наименование объекта!")
            continue

        if gen_inv and not values["-OBJECT_ADDRESS-"]:
            sg.popup_error("Укажите адрес объекта!")
            continue

        if gen_passp and not values["-PASSP_OUTPUT-"]:
            sg.popup_error("Укажите путь сохранения паспортов!")
            continue

        # ---------- Чтение Excel ----------

        print("=" * 60)
        print("Чтение Excel...")

        df, load_error = load_source_data(values["-EXCEL-"])

        if load_error:
            sg.popup_error(load_error)
            continue

        print("Excel успешно загружен.\n")

        try:
            selected_df = select_record_range(
                df,
                start_row,
                end_row,
            )
        except RecordSelectionError as exc:
            sg.popup_error(str(exc))
            continue

        selected_start_row = 1
        selected_end_row = len(selected_df)

        # ---------- Генерация инвентаризации ----------

        if gen_inv:
            print("Создание инвентаризационного отчёта...")

            selected_records = selected_df.to_dict(orient="records")
            inventory_context = InventoryContext(
                object_name=values["-OBJECT_NAME-"].strip(),
                object_address=values["-OBJECT_ADDRESS-"].strip(),
                staff_count=calculate_staff_count(selected_records),
            )

            success, message = generate_inventory_report(
                df=selected_df,
                context=inventory_context,
                output_path=values["-INV_OUTPUT-"],
                start_row=selected_start_row,
                end_row=selected_end_row
            )

            print(message)

            if not success:
                generation_succeeded = False
                sg.popup_error(message)

        # ---------- Генерация паспортов ----------

        if gen_passp:
            print("Создание паспортов оборудования...")

            workflow = get_workflow(values["-DEPARTMENT-"])

            success, message = generate_passports(
                df=selected_df,
                workflow=workflow,
                output_path=values["-PASSP_OUTPUT-"],
                start_row=selected_start_row,
                end_row=selected_end_row
            )

            print(message)

            if not success:
                generation_succeeded = False
                sg.popup_error(message)

        print("\nГенерация завершена.")
        print("=" * 60)

        if generation_succeeded:
            sg.popup_ok("Документы успешно сформированы.")

    window.close()
