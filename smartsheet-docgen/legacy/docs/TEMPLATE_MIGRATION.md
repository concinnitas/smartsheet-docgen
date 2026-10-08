## Word Template Migration

This project now expects a `docxtpl`-compatible `.docx` template.

### 1. Replace legacy bracket tokens

Convert plain-text placeholders like these:

- `[PROJECT_NAME]`
- `[ADDRESS]`
- `[ISSUE_ID]`
- `[LOCATION]`

Into `docxtpl` tags:

- `{{ PROJECT_NAME }}`
- `{{ ADDRESS }}`
- `{{ row.ISSUE_ID }}`
- `{{ row.LOCATION }}`

### 2. Use top-level tags for static fields

These values come from `static_cell_mappings` and should be used directly:

```text
{{ PROJECT_NAME }}
{{ ADDRESS }}
{{ RACKING_TYPE }}
{{ RACKING_LOCATION }}
{{ DC_CAPACITY }}
{{ AC_CAPACITY }}
{{ INV_1_COUNT }}
{{ INV_1_TYPE }}
{{ INV_2_COUNT }}
{{ INV_2_TYPE }}
{{ MOD_COUNT }}
{{ MOD_TYPE }}
{{ MLPE_COUNT }}
{{ MLPE_TYPE }}
{{ IX_TYPE }}
{{ IX_VOLTAGE }}
```

### 3. Wrap repeated issue content in a loop

Anything that should repeat once per grouped Smartsheet row should live inside a `rows` loop:

```text
{% for row in rows %}
Item #{{ row.ISSUE_ID }} - {{ row.LOCATION }}
Status: {{ row.STATUS }}
Submission Date: {{ row.DATE }}
Inspection Comment: {{ row.INSPECTION_COMMENT }}
Contractor Response: {{ row.CONTRACTOR_RESPONSE }}
AGI Comment: {{ row.AGI_COMMENT }}
{% endfor %}
```

### 4. One page per repeated item

If each repeated item should start on a new page, place a page break inside the loop in Word.

Recommended structure:

```text
{% for row in rows %}
{% if not loop.first %}
[manual Word page break here]
{% endif %}

... repeated item content ...
{% endfor %}
```

The simplest way to do this in Word is:

1. Put `{% if not loop.first %}` on its own line.
2. Insert a manual page break on the next line with `Ctrl+Enter`.
3. Put `{% endif %}` on the line after that.

### 5. Keep each tag in one typing action

In Word, do not build tags by editing part of a tag later if you can avoid it. `docxtpl` is most reliable when each tag is typed as plain text in one run, for example:

- Good: `{{ PROJECT_NAME }}`
- Risky: typing `{{ PROJ`, then later inserting `ECT_NAME }}`

### 6. Available repeated-row fields

Current repeated fields from `mappings.json`:

- `row.ISSUE_ID`
- `row.LOCATION`
- `row.INSPECTION_COMMENT`
- `row.STATUS`
- `row.DATE`
- `row.CONTRACTOR_RESPONSE`
- `row.AGI_COMMENT`

### 7. Built-in convenience fields

The renderer also provides:

- `{{ SELECTED_NAME }}`
- `{{ ITEM_COUNT }}`
- `{{ GENERATED_AT }}`

### 8. Important mapping note

The script now prefers column names over raw column indexes for repeated-row mappings. If the Smartsheet column titles differ from the names in `mappings.json`, update the JSON to match the exact Smartsheet column titles.
