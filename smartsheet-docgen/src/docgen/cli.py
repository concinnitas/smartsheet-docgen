"""Command-line entry point.

Examples (from the project folder):
    python run.py --list-sheets
    python run.py --sheet "FCI Tracker" --list-groups
    python run.py --sheet "FCI Tracker" --group "Sunny Acres" --out "output/Sunny Acres FCI.docx"
    python run.py --sheet-id 1234567890 --group "Sunny Acres" --no-photos
"""
import argparse
import sys
from pathlib import Path

from . import settings, smartsheet_io
from .context import SheetView, build_context, group_rows
from .photos import classify
from .render import image_factory_for, render

DEFAULT_TEMPLATE = settings.TEMPLATES_DIR / "FCI_TEMPLATE_docxtpl.docx"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Generate FCI Word reports from Smartsheet.")
    src = ap.add_mutually_exclusive_group()
    src.add_argument("--sheet", help="Sheet name")
    src.add_argument("--sheet-id", type=int, help="Sheet ID (skips the sheet list call)")
    ap.add_argument("--list-sheets", action="store_true")
    ap.add_argument("--list-groups", action="store_true")
    ap.add_argument("--group", help="Value of the group-by column (e.g. project name)")
    ap.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    ap.add_argument("--out", type=Path)
    ap.add_argument("--mappings", type=Path, default=settings.CONFIG_DIR / "mappings.json")
    ap.add_argument("--no-photos", action="store_true", help="Skip attachment download")
    args = ap.parse_args(argv)

    mappings = settings.load_mappings(args.mappings)
    c = smartsheet_io.client()

    if args.list_sheets:
        for name, sid in sorted(smartsheet_io.list_sheets(c).items()):
            print(f"{sid}\t{name}")
        return 0

    sheet_id = args.sheet_id
    if sheet_id is None:
        if not args.sheet:
            ap.error("--sheet or --sheet-id is required")
        sheets = smartsheet_io.list_sheets(c)
        if args.sheet not in sheets:
            ap.error(f"Sheet '{args.sheet}' not found. Use --list-sheets.")
        sheet_id = sheets[args.sheet]
    sheet = smartsheet_io.get_sheet(c, sheet_id)

    if args.list_groups:
        group_by, groups = group_rows(SheetView(sheet), mappings)
        print(f"Groups in column '{group_by}':")
        for g, rows in sorted(groups.items()):
            print(f"  {g}  ({len(rows)} rows)")
        return 0

    if not args.group:
        ap.error("--group is required (use --list-groups to see options)")
    out = args.out or settings.OUTPUT_DIR / f"{args.group} FCI Report.docx"

    img_cfg = mappings.get("image_mappings", {})

    def photo_provider(row):
        photos = smartsheet_io.row_photos(c, sheet_id, row.id)
        before, after, notes = classify(
            photos, img_cfg.get("inspector_emails", []), img_cfg.get("pick", "latest")
        )
        paths = []
        for p in (before, after):
            try:
                paths.append(smartsheet_io.download_photo(c, sheet_id, p) if p else None)
            except Exception as exc:
                notes.append(f"could not use {p.name}: {exc}")
                paths.append(None)
        return paths[0], paths[1], notes

    warnings = []

    def make_context(tpl):
        ctx, w = build_context(
            sheet, args.group, mappings,
            photo_provider=None if args.no_photos else photo_provider,
            image_factory=image_factory_for(tpl, img_cfg.get("width_inches", 3.0)),
        )
        warnings.extend(w)
        return ctx

    path = render(args.template, make_context, out)
    for w in warnings:
        print("  note:", w)
    print(f"Wrote {path}  ({len(warnings)} notes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
