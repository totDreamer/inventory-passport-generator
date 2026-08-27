"""Content-row detection and strict user-range validation."""

import pandas as pd


NON_CONTENT_FIELDS = frozenset({"passport_number", "mac_addr"})


class RecordSelectionError(ValueError):
    """Raised when a requested record range cannot be selected safely."""


def _has_value(value: object) -> bool:
    """Return whether a normalized cell contains meaningful data."""

    if value is None or pd.isna(value):
        return False

    return bool(str(value).strip())


def get_content_records(df):
    """Return content-bearing records in source order with a clean index."""

    content_fields = [
        field
        for field in df.columns
        if field not in NON_CONTENT_FIELDS
    ]

    if not content_fields:
        return df.iloc[0:0].copy().reset_index(drop=True)

    content_mask = df.loc[:, content_fields].apply(
        lambda row: any(_has_value(value) for value in row),
        axis=1,
    )
    return df.loc[content_mask].copy().reset_index(drop=True)


def parse_row_range(start_value, end_value) -> tuple[int, int]:
    """Parse and validate the order-independent parts of a user range."""

    try:
        start_row = int(start_value)
        end_row = int(end_value)
    except (TypeError, ValueError) as exc:
        raise RecordSelectionError(
            "Номера строк должны быть целыми числами."
        ) from exc

    if start_row < 1:
        raise RecordSelectionError(
            "Начальная строка должна быть не меньше 1."
        )

    if end_row < start_row:
        raise RecordSelectionError(
            "Конечная строка не может быть меньше начальной."
        )

    return start_row, end_row


def select_record_range(df, start_row: int, end_row: int):
    """Select a validated inclusive range from content-bearing records."""

    content_records = get_content_records(df)
    available_count = len(content_records)

    if available_count == 0:
        raise RecordSelectionError(
            "В Excel-файле нет содержательных записей."
        )

    if start_row > available_count or end_row > available_count:
        raise RecordSelectionError(
            f"Доступно записей: {available_count}. "
            f"Указан диапазон {start_row}–{end_row}."
        )

    return content_records.iloc[start_row - 1:end_row].copy().reset_index(
        drop=True
    )
