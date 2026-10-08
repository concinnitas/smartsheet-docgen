# smartsheet-docgen (dev workspace)

Rebuild of the Smartsheet -> Word (FCI report) generator.

## Layout

| Path | Purpose |
|---|---|
| `legacy/scripts/` | Copies of V3-V8 and v9s for review (originals untouched in parent folder) |
| `legacy/docs/` | Prior docxtpl migration notes |
| `config/mappings.json` | Column -> placeholder mappings (copied from v9s workspace) |
| `config/.env` | Smartsheet token (created by setup; git-ignored) |
| `templates/` | `FCI AUTOMATION TEMPLATE.docx`, `barebone template.docx` |
| `src/docgen/` | New code: `settings.py`, `smartsheet_io.py`, `render.py` (stubs) |
| `tests/` | Offline smoke test (text + loop + inline image) |
| `output/`, `cache/attachments/` | Generated reports, downloaded row images |

## Setup (Windows)

```powershell
cd "smartsheet-docgen"
.\setup.ps1
```

Creates `.venv`, installs `requirements.txt`, creates `config\.env`, runs tests.
Requires Python 3.9+ (`py` launcher).

See `REVIEW_NOTES.md` for known issues to address during review.
