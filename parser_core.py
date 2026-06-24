"""Ядро парсера — без GUI, для сервера, бота и Web App."""

from __future__ import annotations

import asyncio
import glob
import json
import os
import re
import xml.etree.ElementTree as ET
import zipfile
from collections import deque
from datetime import datetime
from typing import Any, Callable, Awaitable

AioCallback = Callable[[], Awaitable[None]]


def entity_name(e) -> str:
    if e is None:
        return "Неизвестно"
    if t := getattr(e, "title", None):
        return t
    first = getattr(e, "first_name", "") or ""
    last = getattr(e, "last_name", "") or ""
    name = f"{first} {last}".strip()
    if u := getattr(e, "username", None):
        name += f" (@{u})"
    return name or str(getattr(e, "id", "?"))


def chat_title(c) -> str:
    if c is None:
        return "Неизвестный чат"
    return (
        getattr(c, "title", None)
        or getattr(c, "first_name", None)
        or str(getattr(c, "id", "?"))
    )


def normalize_username(value: str) -> str:
    value = (value or "").strip()
    if value.startswith("@"):
        value = value[1:]
    value = value.split()[0] if value else ""
    if not value or value.lower() in {"нет username", "-", "—"}:
        return ""
    return value


def col_letters(cell_ref: str) -> str:
    m = re.match(r"([A-Z]+)", cell_ref or "")
    return m.group(1) if m else ""


def col_index(col: str) -> int:
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def read_xlsx_sheet(path: str) -> dict[tuple[int, int], str]:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    shared: list[str] = []
    cells: dict[tuple[int, int], str] = {}

    with zipfile.ZipFile(path) as zf:
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(".//m:si", ns):
                parts = [t.text or "" for t in si.findall(".//m:t", ns)]
                shared.append("".join(parts))

        sheet_name = next((n for n in zf.namelist() if n.startswith("xl/worksheets/sheet")), None)
        if not sheet_name:
            raise ValueError("В файле Excel не найден лист")

        sheet = ET.fromstring(zf.read(sheet_name))
        for row in sheet.findall(".//m:sheetData/m:row", ns):
            row_idx = int(row.get("r", "0")) - 1
            for cell in row.findall("m:c", ns):
                ref = cell.get("r", "")
                col = col_letters(ref)
                if not col:
                    continue
                col_idx = col_index(col)
                val_node = cell.find("m:v", ns)
                if val_node is None or val_node.text is None:
                    continue
                val = val_node.text
                if cell.get("t") == "s":
                    val = shared[int(val)]
                cells[(row_idx, col_idx)] = str(val).strip()
    return cells


def parse_leads_xlsx(path: str) -> list[dict]:
    cells = read_xlsx_sheet(path)
    if not cells:
        return []

    header_row = 1
    categories: list[tuple[int, str]] = []
    max_col = max(c for _, c in cells)
    for col in range(max_col + 1):
        title = cells.get((header_row, col), "")
        if title:
            categories.append((col, title))

    if not categories:
        raise ValueError("Не найдена строка категорий (ожидается вторая строка файла)")

    leads: list[dict] = []
    seen: set[str] = set()
    max_row = max(r for r, _ in cells)

    for row in range(header_row + 1, max_row + 1):
        for col, category in categories:
            username_raw = cells.get((row, col), "")
            if not username_raw:
                continue
            username = normalize_username(username_raw)
            if not username:
                continue
            key = username.lower()
            if key in seen:
                continue
            seen.add(key)
            description = cells.get((row, col + 1), "")
            leads.append({
                "username": username,
                "name": f"@{username}",
                "chat": category,
                "message": description[:300],
                "source": "import",
            })
    return leads


def parse_manual_usernames(text: str) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for line in text.splitlines():
        chunk = line.replace(";", ",")
        for part in chunk.split(","):
            username = normalize_username(part)
            if not username:
                continue
            key = username.lower()
            if key in seen:
                continue
            seen.add(key)
            found.append(username)
    return found


def make_lead_entry(
    *,
    username: str,
    name: str = "",
    chat: str = "",
    message: str = "",
    source: str = "manual",
    user_id=None,
) -> dict | None:
    username = normalize_username(username)
    if not username and not user_id:
        return None
    return {
        "_key": f"tg_{user_id}" if user_id else f"u_{username.lower()}",
        "id": user_id,
        "name": name or (f"@{username}" if username else "Неизвестно"),
        "username": username or None,
        "chat": chat or ("Ручной ввод" if source == "manual" else "Импорт Excel"),
        "message": (message or "")[:300],
        "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "selected": True,
        "source": source,
    }


class ParserService:
    """Асинхронный сервис парсера для серверного режима."""

    def __init__(self, data_dir: str, on_change: AioCallback | None = None):
        self.data_dir = data_dir
        os.makedirs(data_dir, exist_ok=True)
        self.config_path = os.path.join(data_dir, "config.json")
        self.session_base = os.path.join(data_dir, "tg_session")
        self.results_path = os.path.join(data_dir, "results.txt")
        self.settings_path = os.path.join(data_dir, "settings.json")

        self.client = None
        self.matched_users: list[dict] = []
        self.result_log: deque[str] = deque(maxlen=1000)
        self.mail_log: deque[str] = deque(maxlen=500)

        self.status = "Не подключено"
        self.scan_progress = ""
        self.mail_progress = ""
        self.is_monitoring = False
        self.is_scanning = False
        self.is_mailing = False
        self._keywords: list[str] = []
        self._active_handler = None
        self._phone: str | None = None
        self.auth_state: str | None = None  # code | 2fa
        self.auth_hint = ""
        self.connected_user: dict | None = None

        self.settings = {
            "keywords": "",
            "scan_history": True,
            "history_limit": "0",
            "mail_delay": "3",
            "mail_message": "",
        }
        self._on_change = on_change
        self._load_settings()

    async def _notify(self):
        if self._on_change:
            await self._on_change()

    def _load_settings(self):
        if not os.path.exists(self.settings_path):
            return
        try:
            with open(self.settings_path, encoding="utf-8") as fh:
                data = json.load(fh)
            self.settings.update(data)
        except Exception:
            pass

    def _save_settings(self):
        with open(self.settings_path, "w", encoding="utf-8") as fh:
            json.dump(self.settings, fh, ensure_ascii=False, indent=2)

    def load_config(self) -> dict:
        if not os.path.exists(self.config_path):
            return {"api_id": "", "api_hash": "", "phone": ""}
        try:
            with open(self.config_path, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return {"api_id": "", "api_hash": "", "phone": ""}

    def save_config(self, cfg: dict):
        with open(self.config_path, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, ensure_ascii=False, indent=2)

    def _append_result(self, line: str):
        self.result_log.append(line.rstrip("\n"))
        with open(self.results_path, "a", encoding="utf-8") as fh:
            fh.write(line if line.endswith("\n") else line + "\n")

    def _append_mail_log(self, line: str):
        self.mail_log.append(line.rstrip("\n"))

    async def get_state(self) -> dict:
        cfg = self.load_config()
        return {
            "status": self.status,
            "scan_progress": self.scan_progress,
            "mail_progress": self.mail_progress,
            "connected": self.client is not None and self.connected_user is not None,
            "connected_user": self.connected_user,
            "auth_state": self.auth_state,
            "auth_hint": self.auth_hint,
            "is_monitoring": self.is_monitoring,
            "is_scanning": self.is_scanning,
            "is_mailing": self.is_mailing,
            "keywords": self.settings.get("keywords", ""),
            "scan_history": self.settings.get("scan_history", True),
            "history_limit": self.settings.get("history_limit", "0"),
            "mail_delay": self.settings.get("mail_delay", "3"),
            "mail_message": self.settings.get("mail_message", ""),
            "leads_count": len(self.matched_users),
            "selected_count": len([u for u in self.matched_users if u.get("selected", True)]),
            "config": {
                "api_id": cfg.get("api_id", ""),
                "api_hash": "••••••" if cfg.get("api_hash") else "",
                "phone": cfg.get("phone", ""),
            },
            "results_tail": list(self.result_log)[-30:],
            "mail_log_tail": list(self.mail_log)[-20:],
        }

    def get_leads(self) -> list[dict]:
        return list(self.matched_users)

    async def update_settings(self, data: dict):
        for key in ("keywords", "scan_history", "history_limit", "mail_delay", "mail_message"):
            if key in data:
                self.settings[key] = data[key]
        self._save_settings()
        await self._notify()

    async def set_api_config(self, api_id: str, api_hash: str, phone: str):
        self.save_config({
            "api_id": str(api_id).strip(),
            "api_hash": api_hash.strip(),
            "phone": phone.strip(),
        })
        await self._notify()

    async def reset_session(self):
        await self.disconnect()
        for pattern in (
            self.session_base,
            self.session_base + ".session",
            self.session_base + ".session-journal",
        ):
            for path in glob.glob(pattern + "*"):
                try:
                    os.remove(path)
                except OSError:
                    pass
        self.auth_state = None
        self.auth_hint = ""
        self.connected_user = None
        self.status = "Сессия сброшена"
        await self._notify()

    async def reset_all(self):
        await self.reset_session()
        if os.path.exists(self.config_path):
            os.remove(self.config_path)
        self.status = "API и сессия сброшены"
        await self._notify()

    async def connect(self):
        cfg = self.load_config()
        api_id_str = str(cfg.get("api_id", "")).strip()
        api_hash = str(cfg.get("api_hash", "")).strip()
        phone = str(cfg.get("phone", "")).strip()
        if not all([api_id_str, api_hash, phone]):
            raise ValueError("Заполните API ID, API Hash и телефон")
        try:
            api_id = int(api_id_str)
        except ValueError as exc:
            raise ValueError("API ID должен быть числом") from exc

        from telethon import TelegramClient
        from telethon.errors import FloodWaitError

        self._phone = phone
        self.status = "Подключение…"
        self.auth_state = None
        await self._notify()

        if self.client:
            try:
                await self.client.disconnect()
            except Exception:
                pass

        self.client = TelegramClient(self.session_base, api_id, api_hash)
        await self.client.connect()

        if not await self.client.is_user_authorized():
            try:
                sent = await self.client.send_code_request(phone)
                code_type = type(sent.type).__name__
                self.auth_hint = (
                    "Код в приложении Telegram"
                    if "App" in code_type
                    else "Код отправлен по SMS"
                )
            except FloodWaitError as e:
                self.status = f"FloodWait: подождите {e.seconds} сек."
                await self._notify()
                raise ValueError(self.status) from e
            self.auth_state = "code"
            self.status = self.auth_hint
        else:
            me = await self.client.get_me()
            await self._on_connected(me)
        await self._notify()

    async def resend_code(self):
        if not self.client or not self._phone:
            raise ValueError("Сначала начните подключение")
        from telethon.errors import FloodWaitError
        try:
            sent = await self.client.send_code_request(self._phone)
            code_type = type(sent.type).__name__
            self.auth_hint = (
                "Новый код в Telegram"
                if "App" in code_type
                else "Новый код по SMS"
            )
            self.status = self.auth_hint
        except FloodWaitError as e:
            self.status = f"FloodWait: {e.seconds} сек."
            raise ValueError(self.status) from e
        await self._notify()

    async def submit_code(self, code: str):
        if not self.client or not self._phone:
            raise ValueError("Нет активного подключения")
        from telethon.errors import SessionPasswordNeededError
        code = code.strip()
        if not code:
            raise ValueError("Введите код")
        try:
            await self.client.sign_in(self._phone, code)
            me = await self.client.get_me()
            await self._on_connected(me)
        except SessionPasswordNeededError:
            self.auth_state = "2fa"
            self.status = "Требуется пароль 2FA"
        except Exception as ex:
            self.status = f"Ошибка кода: {ex}"
            raise
        await self._notify()

    async def submit_2fa(self, password: str):
        if not self.client:
            raise ValueError("Нет активного подключения")
        try:
            await self.client.sign_in(password=password)
            me = await self.client.get_me()
            await self._on_connected(me)
        except Exception as ex:
            self.status = f"Ошибка 2FA: {ex}"
            raise
        await self._notify()

    async def _on_connected(self, me):
        name = f"{me.first_name or ''} {me.last_name or ''}".strip()
        tag = f"@{me.username}" if me.username else me.phone
        self.connected_user = {
            "id": me.id,
            "name": name,
            "username": me.username,
            "phone": me.phone,
            "tag": tag,
        }
        self.auth_state = None
        self.auth_hint = ""
        self.status = f"Подключено: {name} ({tag})"
        await self._notify()

    async def disconnect(self):
        self.stop_monitoring()
        if self.client:
            try:
                await self.client.disconnect()
            except Exception:
                pass
        self.client = None
        self.connected_user = None
        self.auth_state = None
        self.status = "Отключено"
        await self._notify()

    def _parse_keywords(self) -> list[str]:
        kws = [
            " ".join(k.split()).lower()
            for k in self.settings.get("keywords", "").split(",")
            if k.strip()
        ]
        if not kws:
            raise ValueError("Введите ключевые слова")
        return kws

    def _history_limit(self):
        try:
            val = int(str(self.settings.get("history_limit", "0")).strip())
            return None if val <= 0 else val
        except ValueError:
            return None

    def _text_matches(self, raw: str) -> str | None:
        text = " ".join((raw or "").split()).lower()
        if not text or not any(kw in text for kw in self._keywords):
            return None
        return text

    def _format_match_line(self, prefix: str, c_name: str, s_name: str, sender, text: str) -> str:
        uname = f"@{sender.username}" if getattr(sender, "username", None) else "нет username"
        return f"[{prefix}]  Чат: {c_name}  |  От: {s_name}  |  Username: {uname}  |  {text}"

    async def _register_match(self, sender, chat, text: str, msg_date=None, *, source: str = "live"):
        if not hasattr(sender, "id") or not sender.id:
            return

        s_name = entity_name(sender)
        c_name = chat_title(chat)
        if msg_date:
            prefix = msg_date.strftime("%Y-%m-%d %H:%M") if hasattr(msg_date, "strftime") else str(msg_date)
            if source == "history":
                prefix = f"История {prefix}"
        else:
            prefix = datetime.now().strftime("%H:%M:%S")

        entry = {
            "_key": f"tg_{sender.id}",
            "id": sender.id,
            "name": s_name,
            "username": getattr(sender, "username", None),
            "chat": c_name,
            "message": text[:300],
            "date": prefix,
            "selected": True,
            "source": "парсер",
        }

        existing = next((u for u in self.matched_users if u.get("_key") == entry["_key"]), None)
        if existing:
            existing.update({"chat": c_name, "message": text[:300], "date": prefix})
        else:
            self.matched_users.append(entry)

        line = self._format_match_line(prefix, c_name, s_name, sender, text)
        self._append_result(line)
        await self._notify()

    async def start_monitoring(self):
        if not self.client or not self.connected_user:
            raise ValueError("Сначала подключитесь к Telegram")
        self._keywords = self._parse_keywords()
        self.is_monitoring = True
        self.status = f"Мониторинг: {', '.join(self._keywords)}"
        await self._async_add_handler()
        if self.settings.get("scan_history", True):
            asyncio.create_task(self.scan_history())
        await self._notify()

    async def scan_history(self):
        if not self.client:
            raise ValueError("Сначала подключитесь")
        if not self._keywords:
            self._keywords = self._parse_keywords()
        self.is_scanning = True
        self.status = f"Сканирование истории: {', '.join(self._keywords)}"
        await self._notify()
        try:
            dialogs = await self.client.get_dialogs()
            total = len(dialogs)
            limit = self._history_limit()
            found = 0
            for idx, dialog in enumerate(dialogs, 1):
                if not self.is_scanning:
                    break
                chat = dialog.entity
                c_name = chat_title(chat)
                self.scan_progress = f"[{idx}/{total}] {c_name}"
                await self._notify()
                try:
                    async for message in self.client.iter_messages(chat, limit=limit):
                        if not self.is_scanning:
                            break
                        matched = self._text_matches(message.text or message.message or "")
                        if not matched:
                            continue
                        sender = await message.get_sender()
                        if sender is None:
                            continue
                        await self._register_match(sender, chat, matched, message.date, source="history")
                        found += 1
                except Exception as ex:
                    self._append_result(f"[Ошибка «{c_name}»: {ex}]")
            self.scan_progress = f"Готово: {found} совпадений в {total} чатах"
            if self.is_monitoring:
                self.status = f"Мониторинг: {', '.join(self._keywords)}"
            else:
                self.status = self.scan_progress
        finally:
            self.is_scanning = False
            await self._notify()

    async def _async_add_handler(self):
        from telethon import events

        async def on_message(event):
            if not self.is_monitoring:
                return
            matched = self._text_matches(event.message.text or event.message.message or "")
            if not matched:
                return
            try:
                sender = await event.get_sender()
                chat = await event.get_chat()
                await self._register_match(sender, chat, matched, source="live")
            except Exception as ex:
                self._append_result(f"[Ошибка: {ex}]")
                await self._notify()

        self._active_handler = on_message
        self.client.add_event_handler(on_message, events.NewMessage)

    def stop_monitoring(self):
        was_active = self.is_monitoring or self.is_scanning
        self.is_monitoring = False
        self.is_scanning = False
        if self.client and self._active_handler:
            handler = self._active_handler
            self.client.remove_event_handler(handler)
        self._active_handler = None
        if was_active:
            self._save_summary()
            self.scan_progress = ""
            if self.connected_user:
                self.status = "Остановлено"
            else:
                self.status = "Отключено"

    def _save_summary(self):
        if not self.matched_users:
            return
        ts_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sep = "=" * 60
        lines = [
            f"\n{sep}",
            f"СЕССИЯ ЗАВЕРШЕНА: {ts_now}",
            f"Найдено: {len(self.matched_users)}",
            sep,
        ]
        for i, u in enumerate(self.matched_users, 1):
            uname = f"@{u['username']}" if u.get("username") else "—"
            msg = (u.get("message") or "")[:60]
            lines.append(f"  {i}. {u['name']} | {uname} | {u.get('chat', '')} | {msg}")
        lines.append(sep)
        with open(self.results_path, "a", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")

    async def clear_leads(self):
        self.matched_users.clear()
        self.result_log.clear()
        await self._notify()

    def upsert_leads(self, leads: list[dict]) -> tuple[int, int]:
        added = updated = 0
        for raw in leads:
            entry = make_lead_entry(
                username=raw.get("username", ""),
                name=raw.get("name", ""),
                chat=raw.get("chat", ""),
                message=raw.get("message", ""),
                source=raw.get("source", "manual"),
                user_id=raw.get("id"),
            )
            if not entry:
                continue
            existing = next((u for u in self.matched_users if u.get("_key") == entry["_key"]), None)
            if existing:
                existing.update({
                    "name": entry["name"],
                    "chat": entry["chat"],
                    "message": entry["message"],
                    "date": entry["date"],
                    "source": entry["source"],
                    "username": entry["username"],
                })
                updated += 1
            else:
                self.matched_users.append(entry)
                added += 1
        return added, updated

    async def add_manual_usernames(self, text: str) -> tuple[int, int]:
        usernames = parse_manual_usernames(text)
        if not usernames:
            raise ValueError("Нет username для добавления")
        added, updated = self.upsert_leads([{"username": u, "source": "manual"} for u in usernames])
        await self._notify()
        return added, updated

    async def import_xlsx_file(self, path: str) -> tuple[int, int, int]:
        leads = parse_leads_xlsx(path)
        if not leads:
            raise ValueError("В файле не найдено username")
        added, updated = self.upsert_leads(leads)
        await self._notify()
        return len(leads), added, updated

    async def remove_imported_leads(self) -> int:
        before = len(self.matched_users)
        self.matched_users = [u for u in self.matched_users if u.get("source") not in {"manual", "import"}]
        removed = before - len(self.matched_users)
        await self._notify()
        return removed

    async def set_lead_selection(self, keys: list[str] | None, selected: bool):
        if keys is None:
            for u in self.matched_users:
                u["selected"] = selected
        else:
            key_set = set(keys)
            for u in self.matched_users:
                if u.get("_key") in key_set:
                    u["selected"] = selected
        await self._notify()

    async def toggle_lead(self, key: str):
        for u in self.matched_users:
            if u.get("_key") == key:
                u["selected"] = not u.get("selected", True)
                break
        await self._notify()

    async def _resolve_recipient(self, user: dict):
        if user.get("id"):
            return user["id"]
        username = normalize_username(user.get("username", ""))
        if not username:
            raise ValueError("Нет username")
        entity = await self.client.get_entity(username)
        user["id"] = entity.id
        user["_key"] = f"tg_{entity.id}"
        if not user.get("name") or user["name"] == f"@{username}":
            user["name"] = entity_name(entity)
        return entity

    async def send_mailing(self, message: str | None = None, delay: float | None = None):
        if not self.client:
            raise ValueError("Сначала подключитесь")
        recipients = [u for u in self.matched_users if u.get("selected", True)]
        if not recipients:
            raise ValueError("Никто не выбран для рассылки")
        msg = (message or self.settings.get("mail_message", "")).strip()
        if not msg:
            raise ValueError("Введите текст сообщения")
        if delay is None:
            try:
                delay = max(0.5, float(self.settings.get("mail_delay", "3")))
            except ValueError:
                delay = 3.0

        self.settings["mail_message"] = msg
        self._save_settings()
        self.is_mailing = True
        self.status = "Рассылка…"
        await self._notify()

        total = len(recipients)
        success = 0
        try:
            for i, user in enumerate(recipients, 1):
                label = user.get("name") or f"@{user.get('username', '?')}"
                try:
                    target = await self._resolve_recipient(user)
                    await self.client.send_message(target, msg)
                    line = f"[{i}/{total}] ✓ {label}"
                    success += 1
                except Exception as ex:
                    line = f"[{i}/{total}] ✗ {label}: {ex}"
                self._append_mail_log(line)
                self.mail_progress = f"{i}/{total}"
                await self._notify()
                if i < total:
                    await asyncio.sleep(delay)
            summary = f"Рассылка: {success}/{total} успешно"
            self._append_mail_log(summary)
            self.status = summary
        finally:
            self.is_mailing = False
            self.mail_progress = ""
            await self._notify()

        return {"success": success, "total": total}
