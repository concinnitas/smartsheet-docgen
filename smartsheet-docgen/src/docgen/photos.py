"""Before/after photo selection from Smartsheet row attachments.

Rules (see config/mappings.json -> image_mappings):
  1. Only image file attachments are considered (row + row-discussion attachments).
  2. Uploaded by someone in `inspector_emails`  -> BEFORE candidate.
     Uploaded by anyone else                    -> AFTER candidate.
  3. If every photo came from one side and there are 2+, fall back to date:
     oldest -> BEFORE, newest -> AFTER.
  4. When a side has several candidates, `pick` ("latest" | "earliest") decides.
Pure functions only; no API calls here so the rules are unit-testable.
"""
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional, Tuple

WORD_SAFE_EXT = {".jpg", ".jpeg", ".png", ".gif", ".bmp"}


@dataclass
class PhotoInfo:
    id: int
    name: str
    created_at: datetime
    created_by: str  # uploader email, lower-cased ("" if unknown)
    mime_type: str = ""


def is_image(name: str, mime_type: str) -> bool:
    if (mime_type or "").lower().startswith("image/"):
        return True
    return any(name.lower().endswith(ext) for ext in WORD_SAFE_EXT | {".heic", ".webp", ".tif", ".tiff"})


def _choose(items: List[PhotoInfo], pick: str) -> PhotoInfo:
    ordered = sorted(items, key=lambda p: p.created_at)
    return ordered[0] if pick == "earliest" else ordered[-1]


def classify(
    photos: List[PhotoInfo], inspector_emails, pick: str = "latest"
) -> Tuple[Optional[PhotoInfo], Optional[PhotoInfo], List[str]]:
    """Return (before, after, notes)."""
    notes: List[str] = []
    inspectors = {e.strip().lower() for e in inspector_emails if e}
    mine = [p for p in photos if p.created_by in inspectors]
    others = [p for p in photos if p.created_by not in inspectors]

    if mine and others:
        before, after = _choose(mine, pick), _choose(others, pick)
        if len(mine) > 1:
            notes.append(f"{len(mine)} inspector photos; used {pick}: {before.name}")
        if len(others) > 1:
            notes.append(f"{len(others)} contractor photos; used {pick}: {after.name}")
        return before, after, notes

    side = mine or others
    if not side:
        return None, None, notes
    if len(side) == 1:
        return (side[0], None, notes) if mine else (None, side[0], notes)

    ordered = sorted(side, key=lambda p: p.created_at)
    who = "inspector" if mine else "contractor"
    notes.append(
        f"All {len(side)} photos uploaded by {who}; split by date "
        f"(before={ordered[0].name}, after={ordered[-1].name})"
    )
    return ordered[0], ordered[-1], notes
