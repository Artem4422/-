"""Настройки серверного режима."""

from __future__ import annotations

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.getenv("PARSER_DATA_DIR", str(BASE_DIR / "server_data")))
WEB_DIR = BASE_DIR / "web"

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
WEBAPP_URL = os.getenv("WEBAPP_URL", "http://localhost:8080")
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "8080"))

_admin_raw = os.getenv("ADMIN_IDS", "")
ADMIN_IDS: set[int] = {
    int(x.strip()) for x in _admin_raw.split(",") if x.strip().isdigit()
}

API_SECRET = os.getenv("API_SECRET", "")
