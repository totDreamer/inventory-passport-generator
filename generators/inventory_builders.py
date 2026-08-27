"""
Функции подготовки данных для инвентаризационного отчета.
Каждая функция возвращает строку, которая будет записана в ячейку Excel.
"""


# ==========================================================
# Универсальный билдер поля
# ==========================================================

def field(name):
    """
    Возвращает функцию, которая получает значение
    указанного поля из строки DataFrame.
    """

    def getter(data):
        value = getattr(data, name, "")

        if value is None:
            return ""

        return str(value).strip()

    return getter


# ==========================================================
# Вспомогательные функции
# ==========================================================

def clean_value(value):
    """Return a trimmed value without leaking missing-value sentinels."""

    if value is None or value != value:
        return ""

    return str(value).strip()


def format_drive(mark, size):
    """
    Формирует строку накопителя.
    """

    mark = clean_value(mark)
    size = clean_value(size)

    if not mark and not size:
        return ""

    if mark and size:
        return f"{mark} {size}GB"

    return mark or f"{size}GB"


# ==========================================================
# Сложные поля
# ==========================================================

def build_pc(data):
    """
    Производитель + модель ПК.
    """

    return " ".join(
        value
        for value in (
            str(getattr(data, "pc_mark", "")).strip(),
            str(getattr(data, "pc_model", "")).strip(),
        )
        if value
    )


def build_cpu(data):
    """
    Модель процессора + частота.
    """

    model = clean_value(getattr(data, "cpu_model", ""))
    freq = clean_value(getattr(data, "cpu_freq", "")).replace(",", ".")

    if model and freq:
        return f"{model}, {freq}GHz"

    return model or (f"{freq}GHz" if freq else "")


def build_ram(data):
    """
    Объем оперативной памяти.
    """

    memory_type = clean_value(getattr(data, "ddr_type", ""))
    size = clean_value(getattr(data, "ddr_size", ""))

    if memory_type and size:
        return f"{memory_type}, {size}GB"

    return memory_type or (f"{size}GB" if size else "")


def build_storage(data):
    """
    Формирует список накопителей.
    """

    hard_drives = [
        format_drive(
            getattr(data, "hdd_mark", ""),
            getattr(data, "hdd_size", ""),
        ),
        format_drive(
            getattr(data, "hdd2_mark", ""),
            getattr(data, "hdd2_size", ""),
        ),
    ]
    solid_state_drives = [
        format_drive(
            getattr(data, "ssd_mark", ""),
            getattr(data, "ssd_size", ""),
        ),
        format_drive(
            getattr(data, "ssd2_mark", ""),
            getattr(data, "ssd2_size", ""),
        ),
    ]

    hdd_value = ", ".join(device for device in hard_drives if device)
    ssd_value = ", ".join(
        device for device in solid_state_drives if device
    )

    if hdd_value and ssd_value:
        return f"{hdd_value}({ssd_value})"

    if ssd_value:
        return f"({ssd_value})"

    return hdd_value


def build_monitor(data):
    """
    Производитель + модель монитора.
    """

    return " ".join(
        value
        for value in (
            str(getattr(data, "monitor_mark", "")).strip(),
            str(getattr(data, "monitor_model", "")).strip(),
        )
        if value
    )


def build_printer(data):
    """
    Тип + модель принтера.
    """

    return " ".join(
        value
        for value in (
            str(getattr(data, "printer_dev", "")).strip(),
            str(getattr(data, "printer_model", "")).strip(),
        )
        if value
    )


def build_phone(data):
    """
    Тип + модель IP-телефона.
    """

    return " ".join(
        value
        for value in (
            str(getattr(data, "ip_dev", "")).strip(),
            str(getattr(data, "ip_model", "")).strip(),
        )
        if value
    )
