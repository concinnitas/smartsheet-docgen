"""Build the docxtpl render context from a Smartsheet sheet + mappings.json.

Ported from legacy/scripts/v9s.py. Smartsheet objects are only read via
.columns / .rows / .cells / .title / .id / .display_value / .value, so plain
stand-in objects work in tests.
"""
from datetime import date
from typing import Callable, Dict, List, Optional, Tuple

PhotoProvider = Callable[[object], Tuple[Optional[str], Optional[str], List[str]]]


def normalize_key(value) -> str:
    return str(value).strip().casefold()


def tag_name(key: str) -> str:
    """Accepts legacy '[NAME]' keys too."""
    return key.strip().strip("[]").replace("-", "_")


def cell_value(cell):
    if cell is None:
        return None
    for attr in ("display_value", "value"):
        value = getattr(cell, attr, None)
        if isinstance(value, str):
            value = value.strip()
        if value not in (None, ""):
            return value
    return None


def or_fallback(value, fallback):
    return fallback if value in (None, "") else value


class SheetView:
    """Column-title lookup + per-row cell maps."""

    def __init__(self, sheet):
        self.sheet = sheet
        self.columns = {normalize_key(c.title): c for c in sheet.columns}
        self.cells = [{c.column_id: c for c in getattr(r, "cells", [])} for r in sheet.rows]

    def column(self, ref):
        if isinstance(ref, int):
            cols = list(self.sheet.columns)
            return cols[ref] if 0 <= ref < len(cols) else None
        return self.columns.get(normalize_key(ref))

    def value(self, row_index: int, column) -> object:
        return cell_value(self.cells[row_index].get(column.id))

    def titles(self, limit=12) -> str:
        t = [str(c.title) for c in self.sheet.columns]
        return ", ".join(t[:limit]) + (", ..." if len(t) > limit else "")


def column_ref(details: dict):
    for key in ("column_name", "column_title", "column"):
        if details.get(key) not in (None, ""):
            return details[key]
    return None


def group_rows(view: SheetView, mappings: dict) -> Tuple[str, Dict[str, List[int]]]:
    """{group value: [row indexes]} using dynamic_section_mappings.group_by_column."""
    group_by = mappings.get("dynamic_section_mappings", {}).get("group_by_column") or "Name"
    col = view.column(group_by)
    if col is None:
        raise ValueError(f"Group-by column '{group_by}' not found. Columns: {view.titles()}")
    groups: Dict[str, List[int]] = {}
    for i in range(len(view.sheet.rows)):
        v = view.value(i, col)
        if v not in (None, ""):
            groups.setdefault(str(v).strip(), []).append(i)
    return group_by, groups


def static_context(view: SheetView, mappings: dict):
    ctx, warnings = {}, []
    for key, details in mappings.get("static_cell_mappings", {}).items():
        name, fallback = tag_name(key), details.get("fallback", "")
        loc = details.get("location", {})
        value = None
        try:
            row_index = int(loc.get("row", 1)) - 1
        except (TypeError, ValueError, AttributeError):
            row_index = None
        col = view.column(column_ref(loc)) if isinstance(loc, dict) else None
        if row_index is None or not 0 <= row_index < len(view.sheet.rows):
            warnings.append(f"{name}: row {loc!r} invalid; fallback used")
        elif col is None:
            warnings.append(f"{name}: column {column_ref(loc)!r} not found; fallback used")
        else:
            value = view.value(row_index, col)
            if value in (None, ""):
                warnings.append(f"{name}: cell blank; fallback used")
        ctx[name] = or_fallback(value, fallback)
    return ctx, warnings


def row_contexts(view: SheetView, row_indexes: List[int], mappings: dict,
                 photo_provider: Optional[PhotoProvider] = None, image_factory=None):
    """One dict per issue row. image_factory(path) -> InlineImage (bound to the template)."""
    fields = mappings.get("dynamic_section_mappings", {}).get("fields", {})
    img_cfg = mappings.get("image_mappings", {})
    missing = img_cfg.get("missing_text", "No photo provided")
    rows, warnings = [], []

    for n, i in enumerate(row_indexes, start=1):
        rc = {"INDEX": n, "FIG_BEFORE": 2 * n - 1, "FIG_AFTER": 2 * n}
        for key, details in fields.items():
            name = tag_name(key)
            col = view.column(column_ref(details))
            value = None
            if col is None:
                warnings.append(f"{name}: column {column_ref(details)!r} not found; fallback used")
            else:
                value = view.value(i, col)
            rc[name] = or_fallback(value, details.get("fallback", ""))

        rc["PHOTO_BEFORE"] = rc["PHOTO_AFTER"] = missing
        if photo_provider is not None:
            before, after, notes = photo_provider(view.sheet.rows[i])
            label = f"Item {rc.get('ISSUE_ID', n)}"
            warnings.extend(f"{label}: {note}" for note in notes)
            if before and image_factory:
                rc["PHOTO_BEFORE"] = image_factory(before)
            if after and image_factory:
                rc["PHOTO_AFTER"] = image_factory(after)
            if not before:
                warnings.append(f"{label}: no BEFORE photo")
            if not after:
                warnings.append(f"{label}: no AFTER photo")
        rows.append(rc)
    return rows, warnings


def submission_date(d: Optional[date] = None) -> str:
    d = d or date.today()
    return f"{d:%B} {d.day}, {d.year}"


def build_context(sheet, group: str, mappings: dict, photo_provider=None, image_factory=None):
    view = SheetView(sheet)
    _, groups = group_rows(view, mappings)
    if group not in groups:
        raise ValueError(f"No rows found for '{group}'. Options: {sorted(groups)[:20]}")
    static, w1 = static_context(view, mappings)
    rows, w2 = row_contexts(view, groups[group], mappings, photo_provider, image_factory)
    ctx = {
        **static,
        "rows": rows,
        "SELECTED_NAME": group,
        "ITEM_COUNT": len(rows),
        "SUBMISSION_DATE": submission_date(),
    }
    return ctx, w1 + w2
