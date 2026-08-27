"""Adapters from unified source records to workflow-specific contexts."""

from collections.abc import Callable
from typing import Mapping

from workflow_context import normalize_value

EMPTY_DGD_TEMPLATE_FIELDS = (
    "mouse_dev",
    "mouse_model",
    "keyboard_dev",
    "keyboard_model",
)

PassportContextAdapter = Callable[
    [Mapping[str, object]],
    dict[str, str],
]


def normalize_passport_context(
    record: Mapping[str, object],
) -> dict[str, str]:
    """Convert a source record into trimmed string values for a Word template."""

    return {
        str(field): normalize_value(value)
        for field, value in record.items()
    }


def build_dgk_ups_value(context: Mapping[str, str]) -> str:
    """Build DGK's single UPS template value from neutral UPS fields."""

    manufacturer_and_model = " ".join(
        value
        for value in (
            context.get("ibp_dev", ""),
            context.get("ibp_model", ""),
        )
        if value
    )

    parts = [manufacturer_and_model]

    if serial_number := context.get("ibp_sn", ""):
        parts.append(f"S/N: {serial_number}")

    if inventory_number := context.get("ibp_inv_num", ""):
        parts.append(f"IN: {inventory_number}")

    return "; ".join(part for part in parts if part)


def build_dgk_passport_context(
    record: Mapping[str, object],
) -> dict[str, str]:
    """Build a normalized DocxTemplate context for the unchanged DGK template."""

    context = normalize_passport_context(record)
    context["ibp"] = build_dgk_ups_value(context)
    return context


def build_dgd_passport_context(
    record: Mapping[str, object],
) -> dict[str, str]:
    """Build a normalized DocxTemplate context for the unchanged DGD template."""

    context = normalize_passport_context(record)

    for field in EMPTY_DGD_TEMPLATE_FIELDS:
        context[field] = ""

    return context


PASSPORT_CONTEXT_ADAPTERS: dict[str, PassportContextAdapter] = {
    "dgd": build_dgd_passport_context,
    "dgk": build_dgk_passport_context,
}


def get_passport_context_adapter(
    adapter_id: str,
) -> PassportContextAdapter:
    """Resolve a configured passport-context adapter by its stable ID."""

    try:
        return PASSPORT_CONTEXT_ADAPTERS[adapter_id]
    except KeyError as exc:
        raise ValueError(
            f"Неизвестный адаптер контекста паспорта: {adapter_id}"
        ) from exc
