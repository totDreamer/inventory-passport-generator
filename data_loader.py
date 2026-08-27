import warnings

import pandas as pd

from schemas import (
    InputSchema,
    UNIFIED_V1_NORMALIZED_HEADERS,
    UNIFIED_V1_SCHEMA,
    validate_headers,
)


class InputSchemaError(ValueError):
    """Raised when a source workbook does not match its input schema."""


class UnexpectedInputColumnsWarning(UserWarning):
    """Warn that non-contract columns were ignored during normalization."""


def load_excel(path, schema: InputSchema = UNIFIED_V1_SCHEMA):
    """Load and validate a source workbook with a row-six header contract."""

    raw_data = pd.read_excel(path, skiprows=5, header=None, dtype=str)

    if raw_data.empty:
        raise InputSchemaError("В Excel-файле отсутствует строка заголовков.")

    raw_headers = raw_data.iloc[0].fillna("").tolist()
    validation = validate_headers(raw_headers, schema)

    errors = []

    if validation.missing_required:
        errors.append(
            "отсутствуют поля: "
            + ", ".join(validation.missing_required)
        )

    if validation.duplicate_headers:
        errors.append(
            "дублируются поля: "
            + ", ".join(validation.duplicate_headers)
        )

    if validation.malformed_headers:
        errors.append(
            "некорректные заголовки: "
            + ", ".join(validation.malformed_headers)
        )

    if errors:
        raise InputSchemaError(
            "Некорректная схема входного Excel: "
            + "; ".join(errors)
        )

    if validation.unexpected_headers:
        warnings.warn(
            "Дополнительные колонки не входят во входную схему и будут "
            "игнорироваться: "
            + ", ".join(validation.unexpected_headers),
            UnexpectedInputColumnsWarning,
            stacklevel=2,
        )

    df = raw_data.iloc[1:].copy()
    df.columns = [str(header).strip() for header in raw_headers]

    for header in schema.optional_headers:
        if header not in df.columns:
            df[header] = ""

    normalized = df.loc[:, list(schema.input_headers)].copy()

    if schema.id == UNIFIED_V1_SCHEMA.id:
        normalized["mac_addr"] = ""
        normalized = normalized.loc[:, list(UNIFIED_V1_NORMALIZED_HEADERS)]

    return normalized.fillna("").reset_index(drop=True)
