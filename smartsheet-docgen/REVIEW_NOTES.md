# Review notes

## Resolved
- Template converted to docxtpl (`tools/convert_template.py`); typos fixed
  (`CONTRATOR_RESPONSE`, `IX-VOLTAGE`); mapping keys are bare underscore names.
- Root cause of image failures in V4-V6: the template's photos were floating
  (anchored) pictures with separate floating caption text boxes, and the
  scripts swapped images by searching runs. Replaced with an inline 2-column
  table filled by docxtpl `InlineImage`.
- Photos come from Smartsheet row attachments, split before/after by uploader,
  then by date (`src/docgen/photos.py`).
- Sheet/attachment listing pages explicitly instead of `include_all`.
- Token lives in `config/.env`, never written back by the app.

## Open
- Static fields read from sheet row 1 (same as v9s). If one sheet holds
  several projects (grouped by `Name`), static values should come from the
  selected group's first row instead. Confirm sheet layout.
- `inspector_emails` must match the email on your Smartsheet account.
- Per-item heading "Item #N – LOCATION" was added (TOC referenced it but the
  section had no heading). Captions number continuously (FIG 1-2, 3-4, ...).
- Empty second-inverter line prints blank when INV_2 is empty; could be wrapped
  in `{%p if INV_2_COUNT %}`.
- GUI (tkinter in v9s) not ported yet; CLI only.
- Old `dev workspace/config.json` holds a plain-text token: rotate it.
- Not yet run against live Smartsheet; first run will confirm the SDK 4.x
  attachment fields (`created_by.email`, `created_at`).
