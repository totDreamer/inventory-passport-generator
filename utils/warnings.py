import warnings
import re

# Сохраняем оригинальную функцию
original_showwarning = warnings.showwarning

# Свои условия
def custom_showwarning(message, category, filename, lineno, file=None, line=None):
    msg = str(message)
    if (
        re.search(r"pkg_resources is deprecated as an API", msg)
        or re.search(r"Data Validation extension is not supported", msg)
    ):
        return  # Не выводим эти два предупреждения
    return original_showwarning(message, category, filename, lineno, file, line)

warnings.showwarning = custom_showwarning