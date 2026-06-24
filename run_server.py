#!/usr/bin/env python3
"""Запуск серверного режима: API + Telegram Bot + Web App."""

from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path

import uvicorn

BASE = Path(__file__).resolve().parent


def _load_env():
    env_path = BASE / "server.env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        os.environ.setdefault(key.strip(), val.strip())


_load_env()

from parser_core import ParserService  # noqa: E402
from server.api import create_app  # noqa: E402
from server.bot import setup_bot  # noqa: E402
from server.settings import BOT_TOKEN, DATA_DIR, HOST, PORT, WEBAPP_URL  # noqa: E402


async def main():
    if not BOT_TOKEN:
        print("Ошибка: задайте BOT_TOKEN в server.env или переменных окружения")
        sys.exit(1)

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    service = ParserService(str(DATA_DIR))
    app = create_app(service)
    bot, dp = setup_bot(service, BOT_TOKEN)

    config = uvicorn.Config(app, host=HOST, port=PORT, log_level="info")
    server = uvicorn.Server(config)

    print(f"API + Web App: http://{HOST}:{PORT}")
    print(f"Web App URL для бота: {WEBAPP_URL}")
    print(f"Данные парсера: {DATA_DIR}")

    await asyncio.gather(
        server.serve(),
        dp.start_polling(bot),
    )


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        pass
