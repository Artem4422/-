"""Проверка Telegram Web App initData."""

from __future__ import annotations

import hashlib
import hmac
import json
import time
from urllib.parse import parse_qsl


def validate_init_data(init_data: str, bot_token: str, max_age: int = 86400) -> dict | None:
    if not init_data or not bot_token:
        return None
    try:
        pairs = dict(parse_qsl(init_data, keep_blank_values=True))
        received_hash = pairs.pop("hash", "")
        if not received_hash:
            return None

        data_check = "\n".join(f"{k}={v}" for k, v in sorted(pairs.items()))
        secret = hmac.new(b"WebAppData", bot_token.encode(), hashlib.sha256).digest()
        calc = hmac.new(secret, data_check.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(calc, received_hash):
            return None

        auth_date = int(pairs.get("auth_date", "0"))
        if auth_date and time.time() - auth_date > max_age:
            return None

        user_raw = pairs.get("user")
        user = json.loads(user_raw) if user_raw else {}
        return {"user": user, "raw": pairs}
    except Exception:
        return None
