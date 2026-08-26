from generators.inventory_builders import (
    field,
    build_pc,
    build_cpu,
    build_ram,
    build_storage,
    build_monitor,
    build_printer,
    build_phone,
)

REPORT_FIELDS = [

    # Пользователь
    ("B", field("user_name")),
    ("C", field("department")),
    ("D", field("cabinet")),

    # ПК
    ("E", build_pc),
    ("F", field("pc_serial_number")),
    ("G", field("pc_inv_number")),

    # Сеть
    ("H", field("login")),
    ("I", field("domain")),
    ("J", field("ip_addr")),

    # ОС
    ("K", field("os")),

    # Процессор
    ("L", build_cpu),

    # RAM
    ("M", build_ram),

    # Диски
    ("N", build_storage),

    # Монитор
    ("O", build_monitor),
    ("P", field("monitor_sn")),
    ("Q", field("monitor_inv_num")),

    # Принтер
    ("R", build_printer),
    ("S", field("printer_sn")),
    ("T", field("printer_inv_num")),
    ("U", field("printer_color")),

    # Телефон
    ("V", build_phone),
    ("W", field("ip_sn")),
    ("X", field("ip_inv")),

    # Сеть
    ("Y", field("mac_addr")),

    # ПО
    ("Z", field("antivirus")),
    ("AA", field("dlp")),
]