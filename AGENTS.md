# Repository Guidelines

## Project Structure & Module Organization

This desktop tool turns an Excel sheet into an inventory workbook and equipment passports. `main.py` starts the GUI in `gui.py`; `config.py` selects templates; `data_loader.py` loads Excel.

Generation code belongs in `generators/`: `inventory.py` writes reports, `passports.py` renders Word files, and `inventory_fields.py` maps report fields. Shared helpers belong in `utils/`. Committed templates are in `templates/inventory/` and `templates/passports/`.

## Reference Data & Expected Output

`input/` and `output/` hold real data and reference results. They are Git-ignored and must never be edited, regenerated in place, or committed. `input/ДГД/` and `input/ДГК/` contain source workbooks. `output/ДГД/` has the expected inventory report; `output/ДГК/` also has the combined passport document. Treat them as the acceptance baseline. No DGD passport-output reference exists.

Source workbooks use five descriptive rows followed by identifiers on row 6; the loader skips the first five. DGD and DGK use one identical 52-column input schema, including `user_name`, `department`, `pc_mark`, `pc_model`, `monitor_model`, `pc_type`, `office_ver`, and canonical UPS fields: `ibp_dev`, `ibp_model`, `ibp_sn`, `ibp_inv_num`. After validation, the loader appends internal `mac_addr`, so its normalized DataFrame has 53 columns. Do not alter either contract without approval.

The Word templates remain unchanged. Workflow adapters pass the canonical `ibp_*` fields to DGD and build DGK's single `ibp` value from them. The source intentionally has no mouse or keyboard fields; adapters must provide empty values for those template variables. `mac_addr` is also absent from the source schema and must default to an empty value when needed.

## Build, Test, and Development Commands

- `python -m pip install -r requirements.txt` installs the runtime dependencies.
- `python main.py` launches the GUI for local development.
- `python -m compileall main.py gui.py config.py data_loader.py generators utils` performs a quick syntax/import compilation check.

There is no automated test suite or build system. Smoke-test a small row range and confirm the `.xlsx` and `.docx` outputs open correctly.

## Coding Style & Naming Conventions

Use four-space indentation, `snake_case` functions and variables, and `UPPER_SNAKE_CASE` constants (for example, `INVENTORY_TEMPLATE`). Keep GUI keys hyphenated and uppercase, such as `"-EXCEL-"`. Prefer small functions that isolate data selection, template rendering, and workbook manipulation. No formatter or linter is configured; avoid unrelated formatting.

## Testing Guidelines

Cover first/last selected rows, an empty cell, and multiple passports. Compare with the applicable ignored `output/` reference without modifying it; check merged cells, formatting, page breaks, and workbook headers. Never overwrite committed templates.

## Commit & Pull Request Guidelines

History uses short, imperative summaries, for example `Add requirements`. Use messages such as `Fix passport page breaks`. Keep commits focused. Pull requests should explain the visible change, manual checks, related issues, and screenshots or output details when relevant.

## Configuration & Template Safety

Add a department through `DEPARTMENTS` in `config.py` and commit its matching template. The current labels are `ДГД / ДГИП / ДВГА` and `ДГК`; use `ДГК`, not `ДК`. Do not commit real inventories, exported passports, or local output paths.
