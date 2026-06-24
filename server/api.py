"""REST API + Web App."""

from __future__ import annotations

import asyncio
import os
import tempfile
from typing import Annotated

from fastapi import Depends, FastAPI, File, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from parser_core import ParserService
from server.auth import validate_init_data
from server.settings import ADMIN_IDS, API_SECRET, BOT_TOKEN, WEB_DIR


class SettingsUpdate(BaseModel):
    keywords: str | None = None
    scan_history: bool | None = None
    history_limit: str | None = None
    mail_delay: str | None = None
    mail_message: str | None = None


class ConfigUpdate(BaseModel):
    api_id: str
    api_hash: str
    phone: str


class CodeBody(BaseModel):
    code: str


class PasswordBody(BaseModel):
    password: str


class ManualLeadsBody(BaseModel):
    text: str


class SelectionBody(BaseModel):
    keys: list[str] | None = None
    selected: bool = True


class ToggleBody(BaseModel):
    key: str


class MailBody(BaseModel):
    message: str | None = None
    delay: float | None = None


def create_app(service: ParserService) -> FastAPI:
    app = FastAPI(title="Telegram Parser Server", version="1.0")
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["*"],
        allow_headers=["*"],
    )

    async def require_admin(
        x_telegram_init_data: Annotated[str | None, Header()] = None,
        x_api_secret: Annotated[str | None, Header()] = None,
    ) -> dict:
        if API_SECRET and x_api_secret == API_SECRET:
            return {"user": {"id": 0, "username": "api"}}
        if not ADMIN_IDS:
            return {"user": {"id": 0, "username": "dev"}}
        data = validate_init_data(x_telegram_init_data or "", BOT_TOKEN)
        if not data:
            raise HTTPException(401, "Unauthorized")
        uid = data["user"].get("id")
        if uid not in ADMIN_IDS:
            raise HTTPException(403, "Forbidden")
        return data

    @app.get("/api/health")
    async def health():
        return {"ok": True}

    @app.get("/api/status")
    async def status(_: dict = Depends(require_admin)):
        return await service.get_state()

    @app.get("/api/leads")
    async def leads(_: dict = Depends(require_admin)):
        return service.get_leads()

    @app.post("/api/settings")
    async def update_settings(body: SettingsUpdate, _: dict = Depends(require_admin)):
        await service.update_settings(body.model_dump(exclude_none=True))
        return await service.get_state()

    @app.get("/api/config")
    async def get_config(_: dict = Depends(require_admin)):
        return service.load_config()

    @app.post("/api/config")
    async def set_config(body: ConfigUpdate, _: dict = Depends(require_admin)):
        await service.set_api_config(body.api_id, body.api_hash, body.phone)
        return {"ok": True}

    @app.post("/api/connect")
    async def connect(_: dict = Depends(require_admin)):
        try:
            await service.connect()
        except ValueError as ex:
            raise HTTPException(400, str(ex)) from ex
        return await service.get_state()

    @app.post("/api/auth/resend")
    async def resend(_: dict = Depends(require_admin)):
        try:
            await service.resend_code()
        except ValueError as ex:
            raise HTTPException(400, str(ex)) from ex
        return await service.get_state()

    @app.post("/api/auth/code")
    async def auth_code(body: CodeBody, _: dict = Depends(require_admin)):
        try:
            await service.submit_code(body.code)
        except Exception as ex:
            raise HTTPException(400, str(ex)) from ex
        return await service.get_state()

    @app.post("/api/auth/2fa")
    async def auth_2fa(body: PasswordBody, _: dict = Depends(require_admin)):
        try:
            await service.submit_2fa(body.password)
        except Exception as ex:
            raise HTTPException(400, str(ex)) from ex
        return await service.get_state()

    @app.post("/api/disconnect")
    async def disconnect(_: dict = Depends(require_admin)):
        await service.disconnect()
        return await service.get_state()

    @app.post("/api/reset/session")
    async def reset_session(_: dict = Depends(require_admin)):
        await service.reset_session()
        return await service.get_state()

    @app.post("/api/reset/all")
    async def reset_all(_: dict = Depends(require_admin)):
        await service.reset_all()
        return await service.get_state()

    @app.post("/api/monitor/start")
    async def monitor_start(_: dict = Depends(require_admin)):
        try:
            await service.start_monitoring()
        except ValueError as ex:
            raise HTTPException(400, str(ex)) from ex
        return await service.get_state()

    @app.post("/api/monitor/stop")
    async def monitor_stop(_: dict = Depends(require_admin)):
        service.stop_monitoring()
        await service._notify()
        return await service.get_state()

    @app.post("/api/monitor/scan")
    async def monitor_scan(_: dict = Depends(require_admin)):
        try:
            asyncio.create_task(service.scan_history())
        except ValueError as ex:
            raise HTTPException(400, str(ex)) from ex
        return await service.get_state()

    @app.post("/api/leads/manual")
    async def leads_manual(body: ManualLeadsBody, _: dict = Depends(require_admin)):
        try:
            added, updated = await service.add_manual_usernames(body.text)
        except ValueError as ex:
            raise HTTPException(400, str(ex)) from ex
        return {"added": added, "updated": updated, "leads": service.get_leads()}

    @app.post("/api/leads/import")
    async def leads_import(file: UploadFile = File(...), _: dict = Depends(require_admin)):
        suffix = ".xlsx" if file.filename and file.filename.endswith(".xlsx") else ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(await file.read())
            tmp_path = tmp.name
        try:
            found, added, updated = await service.import_xlsx_file(tmp_path)
        except ValueError as ex:
            raise HTTPException(400, str(ex)) from ex
        finally:
            os.unlink(tmp_path)
        return {"found": found, "added": added, "updated": updated, "leads": service.get_leads()}

    @app.post("/api/leads/selection")
    async def leads_selection(body: SelectionBody, _: dict = Depends(require_admin)):
        await service.set_lead_selection(body.keys, body.selected)
        return service.get_leads()

    @app.post("/api/leads/toggle")
    async def leads_toggle(body: ToggleBody, _: dict = Depends(require_admin)):
        await service.toggle_lead(body.key)
        return service.get_leads()

    @app.delete("/api/leads/imported")
    async def leads_remove_imported(_: dict = Depends(require_admin)):
        removed = await service.remove_imported_leads()
        return {"removed": removed, "leads": service.get_leads()}

    @app.delete("/api/leads")
    async def leads_clear(_: dict = Depends(require_admin)):
        await service.clear_leads()
        return {"ok": True}

    @app.post("/api/mail/send")
    async def mail_send(body: MailBody, _: dict = Depends(require_admin)):
        try:
            result = await service.send_mailing(body.message, body.delay)
        except ValueError as ex:
            raise HTTPException(400, str(ex)) from ex
        return result

    index_path = WEB_DIR / "index.html"

    @app.get("/")
    async def web_root():
        if index_path.exists():
            return FileResponse(index_path)
        return {"message": "Web App not found"}

    if WEB_DIR.exists():
        app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")

    return app
