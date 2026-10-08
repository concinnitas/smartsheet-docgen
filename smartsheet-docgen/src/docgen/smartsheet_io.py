"""Smartsheet access layer. Placeholder for revised logic after legacy review.

Known API change: unbounded listing (include_all) is sunset; SDK 4.x uses
token pagination (max_items / last_key) on affected list endpoints.
"""
from pathlib import Path

import requests
import smartsheet

from . import settings


def client() -> smartsheet.Smartsheet:
    c = smartsheet.Smartsheet(settings.access_token())
    c.errors_as_exceptions(True)
    return c


def get_sheet(c: smartsheet.Smartsheet, sheet_id: int):
    """Prefer sheet IDs over name lookup to avoid list endpoints entirely."""
    return c.Sheets.get_sheet(sheet_id)


def download_row_images(c, sheet_id: int, row_id: int, dest: Path = None) -> list:
    """Download image attachments on a row; returns local file paths."""
    dest = Path(dest or settings.ATTACHMENT_CACHE) / str(row_id)
    dest.mkdir(parents=True, exist_ok=True)
    paths = []
    for att in c.Attachments.list_row_attachments(sheet_id, row_id).data:
        if att.attachment_type != "FILE" or not (att.mime_type or "").startswith("image/"):
            continue
        url = c.Attachments.get_attachment(sheet_id, att.id).url  # short-lived URL
        out = dest / f"{att.id}_{att.name}"
        if not out.exists():
            r = requests.get(url, timeout=60)
            r.raise_for_status()
            out.write_bytes(r.content)
        paths.append(out)
    return paths
