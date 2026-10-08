"""Smartsheet access layer (SDK 4.x).

Notes on the 2025-2026 API changes:
  * Unbounded listing (include_all) was sunset for folders, workspaces,
    templates, dashboards, webhooks and sharing. Sheets are listed here with
    explicit page/page_size so the call keeps working either way.
  * Attachment download URLs are short-lived; fetch them right before download.
"""
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List

import requests
import smartsheet
from PIL import Image, ImageOps

from . import settings
from .photos import WORD_SAFE_EXT, PhotoInfo, is_image

MAX_IMAGE_PX = 1600  # long edge after normalising; keeps the .docx small


def client() -> smartsheet.Smartsheet:
    c = smartsheet.Smartsheet(settings.access_token())
    c.errors_as_exceptions(True)
    return c


def list_sheets(c) -> Dict[str, int]:
    """Return {sheet name: sheet id}, paging explicitly."""
    sheets, page = {}, 1
    while True:
        result = c.Sheets.list_sheets(page_size=100, page=page)
        for s in result.data:
            sheets[s.name] = s.id
        total_pages = getattr(result, "total_pages", None) or 1
        if page >= total_pages or not result.data:
            return sheets
        page += 1


def get_sheet(c, sheet_id: int):
    return c.Sheets.get_sheet(sheet_id)


def _as_datetime(value) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    return datetime.min.replace(tzinfo=timezone.utc)


def row_photos(c, sheet_id: int, row_id: int) -> List[PhotoInfo]:
    """Image attachments on a row (includes row-discussion attachments)."""
    photos, attachments, page = [], [], 1
    while True:
        result = c.Attachments.list_row_attachments(sheet_id, row_id, page_size=100, page=page)
        attachments.extend(result.data)
        if page >= (getattr(result, "total_pages", None) or 1) or not result.data:
            break
        page += 1
    for att in attachments:
        if str(getattr(att, "attachment_type", "FILE")) != "FILE":
            continue  # links to Google Drive/Box/etc. can't be downloaded here
        name = att.name or f"{att.id}"
        mime = getattr(att, "mime_type", "") or ""
        if not is_image(name, mime):
            continue
        creator = getattr(att, "created_by", None)
        email = (getattr(creator, "email", "") or "").lower()
        photos.append(PhotoInfo(att.id, name, _as_datetime(att.created_at), email, mime))
    return photos


def download_photo(c, sheet_id: int, photo: PhotoInfo, dest_dir: Path = None) -> Path:
    """Download (cached) and normalise to an upright JPEG Word can embed."""
    dest_dir = Path(dest_dir or settings.ATTACHMENT_CACHE)
    dest_dir.mkdir(parents=True, exist_ok=True)
    final = dest_dir / f"{photo.id}.jpg"
    if final.exists():
        return final

    url = c.Attachments.get_attachment(sheet_id, photo.id).url
    raw = dest_dir / f"{photo.id}{Path(photo.name).suffix.lower() or '.bin'}"
    resp = requests.get(url, timeout=120)
    resp.raise_for_status()
    raw.write_bytes(resp.content)

    try:
        with Image.open(raw) as im:
            im = ImageOps.exif_transpose(im)  # phone photos: honour rotation flag
            im.thumbnail((MAX_IMAGE_PX, MAX_IMAGE_PX))
            im.convert("RGB").save(final, "JPEG", quality=85)
    except Exception as exc:  # e.g. HEIC without a plugin
        if raw.suffix in WORD_SAFE_EXT:
            return raw
        raise ValueError(f"Unsupported image format for {photo.name}: {exc}") from exc
    finally:
        if raw != final and raw.exists() and final.exists():
            raw.unlink()
    return final
