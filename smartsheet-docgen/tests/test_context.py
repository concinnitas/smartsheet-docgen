from datetime import date

from docgen.context import build_context, submission_date
from fakes import COLUMNS, ROWS, make_sheet

MAPPINGS = {
    "static_cell_mappings": {
        "PROJECT_NAME": {"location": {"row": 1, "column": "PROJECT NAME"}, "fallback": "?"},
        "[ADDRESS]": {"location": {"row": 1, "column": "ADDRESS"}, "fallback": "?"},
        "MISSING": {"location": {"row": 1, "column": "NOPE"}, "fallback": "fb"},
    },
    "dynamic_section_mappings": {
        "group_by_column": "Name",
        "fields": {
            "ISSUE_ID": {"column_name": "ISSUE ID", "fallback": ""},
            "CONTRACTOR_RESPONSE": {"column_name": "CONTRACTOR RESPONSE", "fallback": "No Contractor Response"},
        },
    },
    "image_mappings": {"missing_text": "No photo provided"},
}


def test_group_rows_static_and_fallbacks():
    ctx, warnings = build_context(make_sheet(COLUMNS, ROWS), "Sunny Acres", MAPPINGS)
    assert ctx["PROJECT_NAME"] == "Sunny Acres Solar" and ctx["ADDRESS"] == "123 Farm Rd"
    assert ctx["MISSING"] == "fb" and any("MISSING" in w for w in warnings)
    assert [r["ISSUE_ID"] for r in ctx["rows"]] == [1, 2]
    assert ctx["rows"][1]["CONTRACTOR_RESPONSE"] == "No Contractor Response"
    assert [(r["FIG_BEFORE"], r["FIG_AFTER"]) for r in ctx["rows"]] == [(1, 2), (3, 4)]
    assert ctx["rows"][0]["PHOTO_BEFORE"] == "No photo provided"


def test_photo_provider_and_factory():
    def provider(row):
        return ("before.jpg", None, ["note"]) if row.id == 1000 else (None, None, [])
    ctx, warnings = build_context(make_sheet(COLUMNS, ROWS), "Sunny Acres", MAPPINGS,
                                  photo_provider=provider, image_factory=lambda p: f"IMG:{p}")
    assert ctx["rows"][0]["PHOTO_BEFORE"] == "IMG:before.jpg"
    assert ctx["rows"][0]["PHOTO_AFTER"] == "No photo provided"
    assert any("no AFTER photo" in w for w in warnings)


def test_submission_date_format():
    assert submission_date(date(2026, 10, 8)) == "October 8, 2026"
