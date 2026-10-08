"""Loads config from config/.env (token) and config/mappings.json."""
import json
import os
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[2]
CONFIG_DIR = ROOT / "config"
TEMPLATES_DIR = ROOT / "templates"
OUTPUT_DIR = ROOT / "output"
ATTACHMENT_CACHE = ROOT / "cache" / "attachments"

load_dotenv(CONFIG_DIR / ".env")


def access_token() -> str:
    token = os.getenv("SMARTSHEET_ACCESS_TOKEN", "").strip()
    if not token:
        raise RuntimeError("SMARTSHEET_ACCESS_TOKEN not set in config/.env")
    return token


def load_mappings(path: Path = CONFIG_DIR / "mappings.json") -> dict:
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)
