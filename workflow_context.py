"""Workflow-level context that does not belong to individual equipment rows."""

from dataclasses import dataclass
from typing import Iterable, Mapping


@dataclass(frozen=True)
class InventoryContext:
    """Values entered once in the GUI for an inventory report."""

    object_name: str
    object_address: str


def normalize_value(value: object) -> str:
    """Return a trimmed string, treating missing values as empty."""

    if value is None:
        return ""

    return str(value).strip()


def is_staff_record(record: Mapping[str, object]) -> bool:
    """Return whether a record represents one employee for staff counting."""

    passport_number = normalize_value(record.get("passport_number"))
    user_name = normalize_value(record.get("user_name"))
    return bool(passport_number) and user_name not in {"", "-"}


def calculate_staff_count(records: Iterable[Mapping[str, object]]) -> int:
    """Count selected records that have a passport number and valid user name."""

    return sum(is_staff_record(record) for record in records)
