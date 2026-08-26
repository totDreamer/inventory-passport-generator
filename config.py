from pathlib import Path

BASE_DIR = Path(__file__).parent

DEPARTMENTS = {
    "ДГД / ДГИП / ДВГА": {
        "passport_template": BASE_DIR / "templates/passports/dgd.docx",
    },

    "ДГК": {
        "passport_template": BASE_DIR / "templates/passports/dgk.docx",
    },
}

INVENTORY_TEMPLATE = (
    BASE_DIR / "templates/inventory/inventory.xlsx"
)


def get_passport_template(department: str):
    """Возвращает путь к шаблону паспортов для подразделения."""
    return DEPARTMENTS[department]["passport_template"]


def get_departments():
    """Возвращает список подразделений для GUI."""
    return list(DEPARTMENTS.keys())