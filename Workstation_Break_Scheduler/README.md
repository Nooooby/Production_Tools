# Workstation & Break Scheduler (Excel VBA)

This project contains the Excel VBA solution for assigning workstations and breaks. It is intentionally separated from other dashboard projects.

## UI Language
All user interface labels, buttons, and messages must be **English**.

## Planned Structure
- `vba/` – VBA modules (main macro + helper modules)
- `templates/` – Excel template files
- `docs/` – usage notes and configuration examples

## Notes
- Keep a single entry macro that runs the full flow: assign stations → insert breaks → validate overlaps → export master.
- Department-specific configuration will live in the Excel input sheets.
