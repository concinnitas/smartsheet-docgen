# smartsheet-docgen

Generates FCI Word reports from a Smartsheet tracker, including before/after
photos pulled from each issue row's attachments.

## Setup (Windows)

```powershell
cd smartsheet-docgen
.\setup.ps1          # creates .venv, installs requirements, creates config\.env, runs tests
```

Put your Smartsheet token in `config\.env` (`SMARTSHEET_ACCESS_TOKEN=...`).

## Usage

```powershell
.\.venv\Scripts\python.exe run.py --list-sheets
.\.venv\Scripts\python.exe run.py --sheet "FCI Tracker" --list-groups
.\.venv\Scripts\python.exe run.py --sheet "FCI Tracker" --group "Sunny Acres"
.\.venv\Scripts\python.exe run.py --sheet-id 1234567890 --group "Sunny Acres" --out "output\Sunny Acres.docx"
```

`--no-photos` skips attachment download. Notes printed after a run list
fallbacks used and photo-selection decisions.

## Photos

Per issue row (row + row-discussion attachments, images only):

| Rule | Result |
|---|---|
| Uploaded by an email in `image_mappings.inspector_emails` | BEFORE (FIG odd: EXAMPLE FIGURE) |
| Uploaded by anyone else | AFTER (FIG even: CORRECTED) |
| All photos from one side, 2+ | oldest = BEFORE, newest = AFTER |
| Several on one side | `pick`: `latest` (default) or `earliest` |
| None | `missing_text` ("No photo provided") |

Photos are rotated upright (phone EXIF), resized to 1600 px, cached in
`cache\attachments\`. HEIC files are skipped with a note.

## Template

`templates\FCI_TEMPLATE_docxtpl.docx` is generated from the original by:

```powershell
.\.venv\Scripts\python.exe tools\convert_template.py "templates\FCI AUTOMATION TEMPLATE.docx" "templates\FCI_TEMPLATE_docxtpl.docx"
```

Tags: `{{ PROJECT_NAME }}` etc. for static fields; inside
`{%p for row in rows %} ... {%p endfor %}`: `{{ row.ISSUE_ID }}`,
`{{ row.PHOTO_BEFORE }}`, `{{ row.FIG_BEFORE }}`, etc. Word asks to update
fields on open; choose Yes to refresh the Table of Contents.

## Layout

| Path | Purpose |
|---|---|
| `src/docgen/` | `cli.py`, `context.py` (sheet -> tags), `photos.py` (before/after rules), `smartsheet_io.py` (API), `render.py` |
| `config/mappings.json` | Column -> tag mappings + `image_mappings` |
| `tools/convert_template.py` | Legacy `[TOKEN]` template -> docxtpl |
| `legacy/` | V3-v9s for reference |
| `tests/` | `pytest -q tests` |
