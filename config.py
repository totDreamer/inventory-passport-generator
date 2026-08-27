from dataclasses import dataclass
from pathlib import Path

BASE_DIR = Path(__file__).parent

UNIFIED_INPUT_SCHEMA = "unified_v1"


@dataclass(frozen=True)
class WorkflowConfig:
    """Configuration for one document-generation workflow."""

    id: str
    label: str
    passport_template: Path
    input_schema: str
    passport_context_adapter: str


INVENTORY_TEMPLATE_PATH = BASE_DIR / "templates/inventory/inventory.xlsx"

WORKFLOWS = {
    "dgd": WorkflowConfig(
        id="dgd",
        label="ДГД / ДГИП / ДВГА",
        passport_template=BASE_DIR / "templates/passports/dgd.docx",
        input_schema=UNIFIED_INPUT_SCHEMA,
        passport_context_adapter="dgd",
    ),
    "dgk": WorkflowConfig(
        id="dgk",
        label="ДГК",
        passport_template=BASE_DIR / "templates/passports/dgk.docx",
        input_schema=UNIFIED_INPUT_SCHEMA,
        passport_context_adapter="dgk",
    ),
}

WORKFLOW_IDS_BY_LABEL = {
    workflow.label: workflow.id
    for workflow in WORKFLOWS.values()
}

def get_workflow(identifier: str) -> WorkflowConfig:
    """Return a workflow by its stable ID or GUI label."""

    workflow_id = WORKFLOW_IDS_BY_LABEL.get(identifier, identifier)
    return WORKFLOWS[workflow_id]


def get_passport_template(identifier: str) -> Path:
    """Return the passport template for a workflow ID or GUI label."""

    return get_workflow(identifier).passport_template


def get_departments() -> list[str]:
    """Return workflow labels for the current GUI selector."""

    return [workflow.label for workflow in WORKFLOWS.values()]
