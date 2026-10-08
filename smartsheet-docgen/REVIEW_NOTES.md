# Review notes (pre-review findings)

## 1. Smartsheet API changes
- Smartsheet sunset unbounded listing (`include_all`) on several list endpoints
  (June 2026); Python SDK 4.0.0 removed those params and moved affected list
  calls to token pagination (`max_items` / `last_key`).
- V3-V8 call `Sheets.list_sheets(include_all=True)` to find sheets by name.
  Likely break point. Fix options: store sheet IDs in config and call
  `get_sheet(id)` directly, or paginate.
- Verify against the changelog before revising:
  https://developers.smartsheet.com/api/smartsheet/changelog

- SDK changelog check: `list_sheets` is NOT among the explicit 4.0.0 removals
  (those are webhooks, sights, workspaces, templates, folders, sharing). So the
  `include_all` call may still work; confirm the actual error with a live run.

## 1b. v9s blockers found in review (independent of API)
- `templates/FCI AUTOMATION TEMPLATE.docx` was never converted: it still has
  24 `[BRACKET]` tokens and zero docxtpl tags, so v9s `validate_template_file`
  rejects it before any Smartsheet data is used.
- Template/mapping mismatches: `[CONTRATOR_RESPONSE]` (typo) vs mapping
  `CONTRACTOR_RESPONSE`; `[IX-VOLTAGE]` (hyphen) vs `IX_VOLTAGE`;
  `[END_OF_SECTION]` is a V6-era section marker that docxtpl doesn't use.
- Token is saved to `config.json` on every fetch/generate (`save_config`).

## 2. Images (main open issue)
- V4-V6 inserted images by locating runs and calling `run.add_picture()` with
  files named `{n}_fig1.jpg` / `{n}_fig2.jpg` from two local folders. Fragile:
  depends on run splitting in Word XML and on file naming/count matching rows.
- v9s moved to docxtpl but has no image handling.
- Proposed approach: docxtpl `InlineImage` objects in the render context,
  placed in the template as `{{ row.PHOTO_1 }}` inside the `{% for row in rows %}`
  loop. Image source either:
  a. Smartsheet row attachments (`list_row_attachments` -> `get_attachment().url`
     -> download to `cache/attachments/<row_id>/`), see `smartsheet_io.download_row_images`; or
  b. Local `FCI Reports/Site Photos/...` folder keyed by an ID column, not row order.
- Decide: source of photos, max photos per item, fixed width vs. per-template size.

## 3. Security
- `dev workspace/config.json` stores the Smartsheet API token in plain text
  inside a OneDrive-synced folder. Recommend rotating the token and using
  `config/.env` (git-ignored) going forward.
