"""Input schema contracts for document-generation workflows."""

from dataclasses import dataclass
from typing import Iterable


@dataclass(frozen=True)
class InputSchema:
    """Immutable contract for normalized source-workbook headers."""

    id: str
    required_headers: tuple[str, ...]
    optional_headers: tuple[str, ...] = ()

    @property
    def input_headers(self) -> tuple[str, ...]:
        """Return canonical input headers in normalized output order."""

        return self.required_headers + self.optional_headers

    @property
    def allowed_headers(self) -> frozenset[str]:
        """Return all headers belonging to this versioned input contract."""

        return frozenset(self.input_headers)


@dataclass(frozen=True)
class SchemaValidation:
    """Header-level validation result without interpreting record values."""

    missing_required: tuple[str, ...]
    duplicate_headers: tuple[str, ...]
    malformed_headers: tuple[str, ...]
    unexpected_headers: tuple[str, ...]

    @property
    def is_valid(self) -> bool:
        """Return whether all required headers are present and unique."""

        return not (
            self.missing_required
            or self.duplicate_headers
            or self.malformed_headers
        )


UNIFIED_V1_INPUT_HEADERS = (
    "passport_number",
    "organisation",
    "department",
    "cabinet",
    "user_name",
    "pc_mark",
    "pc_model",
    "pc_type",
    "pc_serial_number",
    "pc_inv_number",
    "pc_name",
    "login",
    "domain",
    "os",
    "office_ver",
    "ip_addr",
    "cpu_count",
    "cpu_model",
    "cpu_cores",
    "cpu_freq",
    "hdd_mark",
    "hdd_size",
    "hdd2_mark",
    "hdd2_size",
    "ssd_mark",
    "ssd_size",
    "ddr_type",
    "ddr_size",
    "ddr_freq",
    "integrated_vga",
    "discrete_vga",
    "vga_name",
    "vga_size",
    "monitor_mark",
    "monitor_model",
    "monitor_sn",
    "monitor_inv_num",
    "printer_dev",
    "printer_model",
    "printer_color",
    "printer_sn",
    "printer_inv_num",
    "ibp_dev",
    "ibp_model",
    "ibp_sn",
    "ibp_inv_num",
    "ip_dev",
    "ip_model",
    "ip_sn",
    "ip_inv",
    "antivirus",
    "dlp",
)

UNIFIED_V1_REQUIRED_HEADERS = UNIFIED_V1_INPUT_HEADERS
UNIFIED_V1_OPTIONAL_HEADERS: tuple[str, ...] = ()

# Internal fields are added only after the versioned input contract validates.
UNIFIED_V1_NORMALIZED_HEADERS = UNIFIED_V1_INPUT_HEADERS + ("mac_addr",)

UNIFIED_V1_SCHEMA = InputSchema(
    id="unified_v1",
    required_headers=UNIFIED_V1_REQUIRED_HEADERS,
    optional_headers=UNIFIED_V1_OPTIONAL_HEADERS,
)


def validate_headers(
    headers: Iterable[str],
    schema: InputSchema = UNIFIED_V1_SCHEMA,
) -> SchemaValidation:
    """Validate normalized workbook headers against an input schema."""

    normalized = tuple(
        "" if header is None else str(header).strip()
        for header in headers
    )
    unique_headers = frozenset(normalized)
    duplicates = tuple(
        header
        for index, header in enumerate(normalized)
        if header in normalized[:index]
    )

    return SchemaValidation(
        missing_required=tuple(
            header
            for header in schema.required_headers
            if header not in unique_headers
        ),
        duplicate_headers=duplicates,
        malformed_headers=tuple(
            f"колонка {index} не имеет заголовка"
            for index, header in enumerate(normalized, start=1)
            if not header
        ),
        unexpected_headers=tuple(
            header
            for header in normalized
            if header not in schema.allowed_headers
        ),
    )
