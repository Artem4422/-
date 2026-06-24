"""Telegram-бот для управления парсером на сервере."""

from __future__ import annotations

from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
    WebAppInfo,
)

from parser_core import ParserService
from server.settings import ADMIN_IDS, WEBAPP_URL


class Form(StatesGroup):
    api_id = State()
    api_hash = State()
    phone = State()
    code = State()
    password = State()
    keywords = State()
    mail_message = State()
    manual_users = State()


def _is_admin(user_id: int | None) -> bool:
    if not ADMIN_IDS:
        return True
    return user_id in ADMIN_IDS


def main_keyboard() -> InlineKeyboardMarkup:
    rows = [
        [
            InlineKeyboardButton(text="📊 Статус", callback_data="status"),
            InlineKeyboardButton(text="⚙️ API", callback_data="menu_api"),
        ],
        [
            InlineKeyboardButton(text="🔌 Подключение", callback_data="menu_conn"),
            InlineKeyboardButton(text="🔍 Мониторинг", callback_data="menu_mon"),
        ],
        [
            InlineKeyboardButton(text="📨 Рассылка", callback_data="menu_mail"),
            InlineKeyboardButton(text="👥 Лиды", callback_data="leads"),
        ],
        [
            InlineKeyboardButton(
                text="🌐 Web App",
                web_app=WebAppInfo(url=WEBAPP_URL),
            )
        ],
    ]
    return InlineKeyboardMarkup(inline_keyboard=rows)


def back_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="« Меню", callback_data="menu")]
    ])


def api_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="API ID", callback_data="set_api_id")],
        [InlineKeyboardButton(text="API Hash", callback_data="set_api_hash")],
        [InlineKeyboardButton(text="Телефон", callback_data="set_phone")],
        [InlineKeyboardButton(text="📋 Показать конфиг", callback_data="show_config")],
        [InlineKeyboardButton(text="🗑 Сброс сессии", callback_data="reset_session")],
        [InlineKeyboardButton(text="♻️ Сброс API + сессии", callback_data="reset_all")],
        [InlineKeyboardButton(text="« Меню", callback_data="menu")],
    ])


def conn_kb(state: dict) -> InlineKeyboardMarkup:
    rows = []
    if state.get("connected"):
        rows.append([InlineKeyboardButton(text="⏹ Отключить", callback_data="disconnect")])
    else:
        rows.append([InlineKeyboardButton(text="▶ Подключить", callback_data="connect")])
    if state.get("auth_state") == "code":
        rows.append([InlineKeyboardButton(text="🔁 Код снова", callback_data="resend_code")])
        rows.append([InlineKeyboardButton(text="✏️ Ввести код", callback_data="enter_code")])
    if state.get("auth_state") == "2fa":
        rows.append([InlineKeyboardButton(text="🔐 Ввести 2FA", callback_data="enter_2fa")])
    rows.append([InlineKeyboardButton(text="« Меню", callback_data="menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def mon_kb(state: dict) -> InlineKeyboardMarkup:
    rows = [
        [InlineKeyboardButton(text="✏️ Ключевые слова", callback_data="set_keywords")],
    ]
    if state.get("is_monitoring") or state.get("is_scanning"):
        rows.append([InlineKeyboardButton(text="⏹ Остановить", callback_data="mon_stop")])
    else:
        rows.append([InlineKeyboardButton(text="▶ Мониторинг", callback_data="mon_start")])
        rows.append([InlineKeyboardButton(text="🔍 Скан истории", callback_data="mon_scan")])
    rows.append([InlineKeyboardButton(text="« Меню", callback_data="menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def mail_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✏️ Текст рассылки", callback_data="set_mail_msg")],
        [InlineKeyboardButton(text="➕ Username вручную", callback_data="manual_users")],
        [InlineKeyboardButton(text="📤 Отправить выбранным", callback_data="mail_send")],
        [InlineKeyboardButton(text="🗑 Удалить импорт/ручные", callback_data="remove_imported")],
        [InlineKeyboardButton(text="« Меню", callback_data="menu")],
    ])


async def format_status(service: ParserService) -> str:
    s = await service.get_state()
    lines = [
        f"📌 {s['status']}",
        f"Подключено: {'✅' if s['connected'] else '❌'}",
    ]
    if s.get("connected_user"):
        cu = s["connected_user"]
        lines.append(f"Аккаунт: {cu.get('name')} ({cu.get('tag')})")
    if s.get("auth_state"):
        lines.append(f"Авторизация: {s['auth_state']} — {s.get('auth_hint', '')}")
    lines += [
        f"Мониторинг: {'✅' if s['is_monitoring'] else '❌'}",
        f"Сканирование: {'✅' if s['is_scanning'] else '❌'}",
        f"Рассылка: {'✅' if s['is_mailing'] else '❌'}",
        f"Ключевые слова: {s.get('keywords') or '—'}",
        f"Лиды: {s['leads_count']} (выбрано {s['selected_count']})",
    ]
    if s.get("scan_progress"):
        lines.append(f"Прогресс: {s['scan_progress']}")
    if s.get("mail_progress"):
        lines.append(f"Рассылка: {s['mail_progress']}")
    tail = s.get("results_tail") or []
    if tail:
        lines.append("\nПоследние события:")
        lines.extend(tail[-5:])
    return "\n".join(lines)


def setup_bot(service: ParserService, bot_token: str) -> tuple[Bot, Dispatcher]:
    bot = Bot(token=bot_token)
    dp = Dispatcher()
    router = Router()

    @router.message(CommandStart())
    async def cmd_start(message: Message, state: FSMContext):
        if not _is_admin(message.from_user.id if message.from_user else None):
            await message.answer("⛔ Доступ запрещён")
            return
        await state.clear()
        await message.answer(
            "🤖 Управление Telegram Парсером\n\n"
            "Бот управляет парсером на сервере.\n"
            "Полный интерфейс — в Web App.",
            reply_markup=main_keyboard(),
        )

    @router.message(Command("status"))
    async def cmd_status(message: Message):
        if not _is_admin(message.from_user.id if message.from_user else None):
            return
        await message.answer(await format_status(service), reply_markup=main_keyboard())

    @router.callback_query(F.data == "menu")
    async def cb_menu(call: CallbackQuery, state: FSMContext):
        await state.clear()
        await call.message.edit_text("🏠 Главное меню", reply_markup=main_keyboard())
        await call.answer()

    @router.callback_query(F.data == "status")
    async def cb_status(call: CallbackQuery):
        await call.message.edit_text(await format_status(service), reply_markup=main_keyboard())
        await call.answer()

    @router.callback_query(F.data == "menu_api")
    async def cb_menu_api(call: CallbackQuery, state: FSMContext):
        await state.clear()
        await call.message.edit_text("⚙️ Настройки Telegram API", reply_markup=api_kb())
        await call.answer()

    @router.callback_query(F.data == "show_config")
    async def cb_show_config(call: CallbackQuery):
        cfg = service.load_config()
        text = (
            f"API ID: {cfg.get('api_id') or '—'}\n"
            f"API Hash: {'••••••' if cfg.get('api_hash') else '—'}\n"
            f"Телефон: {cfg.get('phone') or '—'}"
        )
        await call.message.edit_text(text, reply_markup=api_kb())
        await call.answer()

    @router.callback_query(F.data == "set_api_id")
    async def cb_set_api_id(call: CallbackQuery, state: FSMContext):
        await state.set_state(Form.api_id)
        await call.message.edit_text("Введите API ID:", reply_markup=back_kb())
        await call.answer()

    @router.callback_query(F.data == "set_api_hash")
    async def cb_set_api_hash(call: CallbackQuery, state: FSMContext):
        await state.set_state(Form.api_hash)
        await call.message.edit_text("Введите API Hash:", reply_markup=back_kb())
        await call.answer()

    @router.callback_query(F.data == "set_phone")
    async def cb_set_phone(call: CallbackQuery, state: FSMContext):
        await state.set_state(Form.phone)
        await call.message.edit_text("Введите телефон (+7…):", reply_markup=back_kb())
        await call.answer()

    @router.message(Form.api_id)
    async def msg_api_id(message: Message, state: FSMContext):
        cfg = service.load_config()
        cfg["api_id"] = message.text.strip()
        service.save_config(cfg)
        await state.clear()
        await message.answer("✅ API ID сохранён", reply_markup=api_kb())

    @router.message(Form.api_hash)
    async def msg_api_hash(message: Message, state: FSMContext):
        cfg = service.load_config()
        cfg["api_hash"] = message.text.strip()
        service.save_config(cfg)
        await state.clear()
        await message.answer("✅ API Hash сохранён", reply_markup=api_kb())

    @router.message(Form.phone)
    async def msg_phone(message: Message, state: FSMContext):
        cfg = service.load_config()
        cfg["phone"] = message.text.strip()
        service.save_config(cfg)
        await state.clear()
        await message.answer("✅ Телефон сохранён", reply_markup=api_kb())

    @router.callback_query(F.data == "reset_session")
    async def cb_reset_session(call: CallbackQuery):
        await service.reset_session()
        await call.message.edit_text("✅ Сессия сброшена", reply_markup=api_kb())
        await call.answer()

    @router.callback_query(F.data == "reset_all")
    async def cb_reset_all(call: CallbackQuery):
        await service.reset_all()
        await call.message.edit_text("✅ API и сессия сброшены", reply_markup=api_kb())
        await call.answer()

    @router.callback_query(F.data == "menu_conn")
    async def cb_menu_conn(call: CallbackQuery, state: FSMContext):
        await state.clear()
        st = await service.get_state()
        await call.message.edit_text(
            f"🔌 Подключение\n\n{st['status']}",
            reply_markup=conn_kb(st),
        )
        await call.answer()

    @router.callback_query(F.data == "connect")
    async def cb_connect(call: CallbackQuery):
        try:
            await service.connect()
            st = await service.get_state()
            await call.message.edit_text(st["status"], reply_markup=conn_kb(st))
        except ValueError as ex:
            await call.answer(str(ex), show_alert=True)
        else:
            await call.answer()

    @router.callback_query(F.data == "disconnect")
    async def cb_disconnect(call: CallbackQuery):
        await service.disconnect()
        st = await service.get_state()
        await call.message.edit_text("Отключено", reply_markup=conn_kb(st))
        await call.answer()

    @router.callback_query(F.data == "resend_code")
    async def cb_resend(call: CallbackQuery):
        try:
            await service.resend_code()
        except ValueError as ex:
            await call.answer(str(ex), show_alert=True)
            return
        st = await service.get_state()
        await call.message.edit_text(st["status"], reply_markup=conn_kb(st))
        await call.answer()

    @router.callback_query(F.data == "enter_code")
    async def cb_enter_code(call: CallbackQuery, state: FSMContext):
        await state.set_state(Form.code)
        await call.message.edit_text("Введите код из Telegram:", reply_markup=back_kb())
        await call.answer()

    @router.callback_query(F.data == "enter_2fa")
    async def cb_enter_2fa(call: CallbackQuery, state: FSMContext):
        await state.set_state(Form.password)
        await call.message.edit_text("Введите пароль 2FA:", reply_markup=back_kb())
        await call.answer()

    @router.message(Form.code)
    async def msg_code(message: Message, state: FSMContext):
        try:
            await service.submit_code(message.text.strip())
        except Exception as ex:
            await message.answer(f"❌ {ex}", reply_markup=back_kb())
            return
        await state.clear()
        st = await service.get_state()
        await message.answer(st["status"], reply_markup=conn_kb(st))

    @router.message(Form.password)
    async def msg_password(message: Message, state: FSMContext):
        try:
            await service.submit_2fa(message.text)
        except Exception as ex:
            await message.answer(f"❌ {ex}", reply_markup=back_kb())
            return
        await state.clear()
        st = await service.get_state()
        await message.answer("✅ Вход выполнен", reply_markup=conn_kb(st))

    @router.callback_query(F.data == "menu_mon")
    async def cb_menu_mon(call: CallbackQuery, state: FSMContext):
        await state.clear()
        st = await service.get_state()
        text = f"🔍 Мониторинг\n\nСлова: {st.get('keywords') or '—'}\n{st['status']}"
        await call.message.edit_text(text, reply_markup=mon_kb(st))
        await call.answer()

    @router.callback_query(F.data == "set_keywords")
    async def cb_set_kw(call: CallbackQuery, state: FSMContext):
        await state.set_state(Form.keywords)
        await call.message.edit_text("Ключевые слова через запятую:", reply_markup=back_kb())
        await call.answer()

    @router.message(Form.keywords)
    async def msg_keywords(message: Message, state: FSMContext):
        await service.update_settings({"keywords": message.text.strip()})
        await state.clear()
        st = await service.get_state()
        await message.answer(f"✅ Слова: {st['keywords']}", reply_markup=mon_kb(st))

    @router.callback_query(F.data == "mon_start")
    async def cb_mon_start(call: CallbackQuery):
        try:
            await service.start_monitoring()
        except ValueError as ex:
            await call.answer(str(ex), show_alert=True)
            return
        st = await service.get_state()
        await call.message.edit_text(st["status"], reply_markup=mon_kb(st))
        await call.answer("Мониторинг запущен")

    @router.callback_query(F.data == "mon_stop")
    async def cb_mon_stop(call: CallbackQuery):
        service.stop_monitoring()
        await service._notify()
        st = await service.get_state()
        await call.message.edit_text("Остановлено", reply_markup=mon_kb(st))
        await call.answer()

    @router.callback_query(F.data == "mon_scan")
    async def cb_mon_scan(call: CallbackQuery):
        import asyncio
        try:
            asyncio.create_task(service.scan_history())
        except ValueError as ex:
            await call.answer(str(ex), show_alert=True)
            return
        await call.answer("Сканирование запущено")

    @router.callback_query(F.data == "menu_mail")
    async def cb_menu_mail(call: CallbackQuery, state: FSMContext):
        await state.clear()
        st = await service.get_state()
        await call.message.edit_text(
            f"📨 Рассылка\n\nВыбрано: {st['selected_count']}/{st['leads_count']}",
            reply_markup=mail_kb(),
        )
        await call.answer()

    @router.callback_query(F.data == "set_mail_msg")
    async def cb_set_mail(call: CallbackQuery, state: FSMContext):
        await state.set_state(Form.mail_message)
        await call.message.edit_text("Текст сообщения для рассылки:", reply_markup=back_kb())
        await call.answer()

    @router.message(Form.mail_message)
    async def msg_mail(message: Message, state: FSMContext):
        await service.update_settings({"mail_message": message.text})
        await state.clear()
        await message.answer("✅ Текст сохранён", reply_markup=mail_kb())

    @router.callback_query(F.data == "manual_users")
    async def cb_manual(call: CallbackQuery, state: FSMContext):
        await state.set_state(Form.manual_users)
        await call.message.edit_text(
            "Username по одному на строку (можно с @):",
            reply_markup=back_kb(),
        )
        await call.answer()

    @router.message(Form.manual_users)
    async def msg_manual(message: Message, state: FSMContext):
        try:
            added, updated = await service.add_manual_usernames(message.text)
        except ValueError as ex:
            await message.answer(str(ex), reply_markup=back_kb())
            return
        await state.clear()
        await message.answer(f"✅ +{added}, обновлено {updated}", reply_markup=mail_kb())

    @router.callback_query(F.data == "mail_send")
    async def cb_mail_send(call: CallbackQuery):
        try:
            result = await service.send_mailing()
        except ValueError as ex:
            await call.answer(str(ex), show_alert=True)
            return
        await call.message.edit_text(
            f"✅ Рассылка: {result['success']}/{result['total']}",
            reply_markup=mail_kb(),
        )
        await call.answer()

    @router.callback_query(F.data == "remove_imported")
    async def cb_remove_imp(call: CallbackQuery):
        removed = await service.remove_imported_leads()
        await call.answer(f"Удалено: {removed}")

    @router.callback_query(F.data == "leads")
    async def cb_leads(call: CallbackQuery):
        leads = service.get_leads()[:20]
        if not leads:
            text = "Список лидов пуст"
        else:
            lines = []
            for u in leads:
                mark = "☑" if u.get("selected", True) else "☐"
                uname = f"@{u['username']}" if u.get("username") else "—"
                lines.append(f"{mark} {u.get('name')} | {uname} | {u.get('chat', '')}")
            text = "\n".join(lines)
            if len(service.get_leads()) > 20:
                text += f"\n\n… всего {len(service.get_leads())}"
        await call.message.edit_text(text[:4000], reply_markup=main_keyboard())
        await call.answer()

    dp.include_router(router)
    return bot, dp
