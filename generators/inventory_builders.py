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

def format_drive(mark, size):
    """
    Формирует строку накопителя.
    """

    mark = str(mark).strip()
    size = str(size).strip()

    if not mark and not size:
        return ""

    if mark and size:
        return f"{mark} {size} GB"

    return mark or f"{size} GB"


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

    model = str(getattr(data, "cpu_model", "")).strip()
    freq = str(getattr(data, "cpu_freq", "")).strip()

    if model and freq:
        return f"{model}, {freq} GHz"

    return model


def build_ram(data):
    """
    Объем оперативной памяти.
    """

    size = str(getattr(data, "ddr_size", "")).strip()

    if not size:
        return ""

    return f"{size} GB"


def build_storage(data):
    """
    Формирует список накопителей.
    """

    devices = [
        format_drive(
            getattr(data, "hdd_mark", ""),
            getattr(data, "hdd_size", ""),
        ),
        format_drive(
            getattr(data, "hdd2_mark", ""),
            getattr(data, "hdd2_size", ""),
        ),
        format_drive(
            getattr(data, "ssd_mark", ""),
            getattr(data, "ssd_size", ""),
        ),
    ]

    return "\n".join(device for device in devices if device)


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