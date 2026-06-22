
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import asyncio
import threading
import json
import os
import sys
from datetime import datetime

# Когда собрано PyInstaller-ом — файлы кладём рядом с .exe, не во временную папку
if getattr(sys, "frozen", False):
    _APP_DIR = os.path.dirname(sys.executable)
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE  = os.path.join(_APP_DIR, "config.json")
RESULTS_FILE = os.path.join(_APP_DIR, "results.txt")
SESSION_FILE = os.path.join(_APP_DIR, "tg_session")


# ── Helpers ─────────────────────────────────────────────────────────────────

def _entity_name(e):
    if e is None:
        return "Неизвестно"
    if t := getattr(e, 'title', None):
        return t
    first = getattr(e, 'first_name', '') or ''
    last  = getattr(e, 'last_name',  '') or ''
    name  = f"{first} {last}".strip()
    if u := getattr(e, 'username', None):
        name += f" (@{u})"
    return name or str(getattr(e, 'id', '?'))

def _chat_title(c):
    if c is None:
        return "Неизвестный чат"
    return (getattr(c, 'title', None)
            or getattr(c, 'first_name', None)
            or str(getattr(c, 'id', '?')))


def _edit_copy(widget, is_text: bool, root: tk.Tk):
    try:
        if is_text:
            if widget.tag_ranges(tk.SEL):
                root.clipboard_clear()
                root.clipboard_append(widget.get(tk.SEL_FIRST, tk.SEL_LAST))
        elif widget.selection_present():
            root.clipboard_clear()
            root.clipboard_append(widget.selection_get())
    except tk.TclError:
        pass


def _edit_paste(widget, is_text: bool, root: tk.Tk):
    try:
        text = root.clipboard_get()
    except tk.TclError:
        return
    try:
        if is_text:
            if str(widget.cget("state")) == tk.DISABLED:
                return
            if widget.tag_ranges(tk.SEL):
                widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
            widget.insert(tk.INSERT, text)
        else:
            if widget.selection_present():
                widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
            widget.insert(tk.INSERT, text)
    except tk.TclError:
        pass


def _edit_cut(widget, is_text: bool, root: tk.Tk):
    _edit_copy(widget, is_text, root)
    try:
        if is_text:
            if str(widget.cget("state")) != tk.DISABLED and widget.tag_ranges(tk.SEL):
                widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
        elif widget.selection_present():
            widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
    except tk.TclError:
        pass


def _edit_select_all(widget, is_text: bool):
    try:
        if is_text:
            widget.tag_add(tk.SEL, "1.0", tk.END)
            widget.mark_set(tk.INSERT, tk.END)
            widget.see(tk.INSERT)
        else:
            widget.select_range(0, tk.END)
            widget.icursor(tk.END)
    except tk.TclError:
        pass


def _bind_edit_menu(root: tk.Tk, widget, *, is_text: bool = False, readonly: bool = False):
    """ПКМ и Cmd+C/V/X/A — для Mac и Windows."""
    menu = tk.Menu(widget, tearoff=0)
    if not readonly:
        menu.add_command(
            label="Вырезать",
            command=lambda: _edit_cut(widget, is_text, root),
        )
    menu.add_command(
        label="Копировать",
        command=lambda: _edit_copy(widget, is_text, root),
    )
    if not readonly:
        menu.add_command(
            label="Вставить",
            command=lambda: _edit_paste(widget, is_text, root),
        )
    menu.add_separator()
    menu.add_command(
        label="Выделить всё",
        command=lambda: _edit_select_all(widget, is_text),
    )

    def show_menu(event):
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    for seq in ("<Button-3>", "<Control-Button-1>", "<Button-2>"):
        widget.bind(seq, show_menu, add=True)


# ── Application ──────────────────────────────────────────────────────────────

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Telegram Парсер")
        self.root.geometry("960x700")
        self.root.minsize(720, 520)

        self.client          = None
        self.is_monitoring   = False
        self._keywords: list = []
        self.matched_users: list = []   # [{id, name, username, chat}]
        self._active_handler = None
        self._ui_queue: list = []

        # Dedicated asyncio loop in background thread
        self.loop = asyncio.new_event_loop()
        threading.Thread(target=self.loop.run_forever, daemon=True).start()

        self._check_telethon()
        self._build_ui()
        self._setup_shortcuts()
        self._load_config()
        self._poll_ui_queue()

    # ── Telethon check ────────────────────────────────────────────────────────

    def _check_telethon(self):
        try:
            import telethon  # noqa: F401
        except ImportError:
            messagebox.showerror(
                "Отсутствует зависимость",
                "Библиотека Telethon не установлена.\n\n"
                "Запустите install.bat или выполните:\n"
                "  pip install telethon"
            )

    # ── Keyboard shortcuts ────────────────────────────────────────────────────

    def _setup_shortcuts(self):
        """Ctrl/Cmd + C/V/X/A по keycode — Mac и Windows, любая раскладка."""
        def on_edit_key(event):
            w  = event.widget
            kc = event.keycode
            is_text = isinstance(w, tk.Text)
            is_entry = isinstance(w, (tk.Entry, ttk.Entry))

            if kc == 67:  # C
                _edit_copy(w, is_text, self.root)
            elif kc == 86 and not (is_text and str(w.cget("state")) == tk.DISABLED):  # V
                _edit_paste(w, is_text, self.root)
            elif kc == 88 and not (is_text and str(w.cget("state")) == tk.DISABLED):  # X
                _edit_cut(w, is_text, self.root)
            elif kc == 65 and (is_text or is_entry):  # A
                _edit_select_all(w, is_text)

        for mod in ("Control", "Command"):
            self.root.bind_all(f"<{mod}-KeyPress>", on_edit_key)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        tab1 = ttk.Frame(nb); nb.add(tab1, text="  Подключение  ")
        tab2 = ttk.Frame(nb); nb.add(tab2, text="  Мониторинг  ")
        tab3 = ttk.Frame(nb); nb.add(tab3, text="  Рассылка  ")

        self._build_conn_tab(tab1)
        self._build_mon_tab(tab2)
        self._build_mail_tab(tab3)

        self.status_var = tk.StringVar(value="Не подключено")
        ttk.Label(self.root, textvariable=self.status_var,
                  relief=tk.SUNKEN, anchor=tk.W,
                  padding=(6, 2)).pack(fill=tk.X, padx=8, pady=(0, 4))

    def _build_conn_tab(self, parent):
        f = ttk.LabelFrame(parent, text="Данные Telegram API", padding=14)
        f.pack(fill=tk.X, padx=24, pady=18)

        fields = [
            ("API ID:",                  "api_id_var",   False),
            ("API Hash:",                "api_hash_var", False),
            ("Номер телефона (+7…):",    "phone_var",    False),
        ]
        for row, (label, attr, secret) in enumerate(fields):
            ttk.Label(f, text=label).grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0, 8))
            var = tk.StringVar(); setattr(self, attr, var)
            kw = {"show": "*"} if secret else {}
            entry = ttk.Entry(f, textvariable=var, width=40, **kw)
            entry.grid(row=row, column=1, pady=5, sticky=tk.W)
            _bind_edit_menu(self.root, entry)

        ttk.Label(
            f,
            text="Получить API ID и Hash: https://my.telegram.org  →  API development tools",
            foreground="#0055cc"
        ).grid(row=3, column=0, columnspan=2, pady=(10, 2), sticky=tk.W)

        ttk.Label(
            f,
            text="Вставка: Cmd+V (Mac) или Ctrl+V (Win)  |  ПКМ по полю — Копировать / Вставить",
            foreground="#666666"
        ).grid(row=4, column=0, columnspan=2, pady=(0, 2), sticky=tk.W)

        bf = ttk.Frame(parent); bf.pack(pady=8)
        self.connect_btn = ttk.Button(bf, text="Подключиться", command=self._connect)
        self.connect_btn.pack(side=tk.LEFT, padx=5)
        self.disconnect_btn = ttk.Button(bf, text="Отключиться",
                                          command=self._disconnect, state=tk.DISABLED)
        self.disconnect_btn.pack(side=tk.LEFT, padx=5)

        # Code frame (hidden until needed)
        self._code_frame = ttk.LabelFrame(parent, text="Код подтверждения", padding=10)
        ttk.Label(self._code_frame, text="Код из Telegram / SMS:").grid(row=0, column=0, padx=(0, 8))
        self.code_var = tk.StringVar()
        code_entry = ttk.Entry(self._code_frame, textvariable=self.code_var, width=14)
        code_entry.grid(row=0, column=1)
        _bind_edit_menu(self.root, code_entry)
        ttk.Button(self._code_frame, text="Подтвердить",
                   command=self._submit_code).grid(row=0, column=2, padx=(8, 0))
        self._resend_btn = ttk.Button(self._code_frame, text="Отправить код снова",
                                      command=self._resend_code)
        self._resend_btn.grid(row=1, column=0, columnspan=3, pady=(8, 0), sticky=tk.W)

        # 2FA frame (hidden until needed)
        self._pw_frame = ttk.LabelFrame(parent, text="Двухфакторная аутентификация (2FA)", padding=10)
        ttk.Label(self._pw_frame, text="Пароль 2FA:").grid(row=0, column=0, padx=(0, 8))
        self.password_var = tk.StringVar()
        pw_entry = ttk.Entry(self._pw_frame, textvariable=self.password_var,
                             show="*", width=24)
        pw_entry.grid(row=0, column=1)
        _bind_edit_menu(self.root, pw_entry)
        ttk.Button(self._pw_frame, text="Войти",
                   command=self._submit_password).grid(row=0, column=2, padx=(8, 0))

    def _build_mon_tab(self, parent):
        kf = ttk.LabelFrame(parent, text="Ключевые слова", padding=10)
        kf.pack(fill=tk.X, padx=16, pady=10)
        ttk.Label(kf, text="Слова через запятую:").pack(side=tk.LEFT)
        self.keywords_var = tk.StringVar()
        kw_entry = ttk.Entry(kf, textvariable=self.keywords_var, width=55)
        kw_entry.pack(side=tk.LEFT, padx=8)
        _bind_edit_menu(self.root, kw_entry)

        bf = ttk.Frame(parent); bf.pack(pady=4)
        self.start_btn = ttk.Button(bf, text="▶  Начать мониторинг",
                                     command=self._start_monitoring, state=tk.DISABLED)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(bf, text="⏹  Остановить",
                                    command=self._stop_monitoring, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Очистить список",
                   command=self._clear_results).pack(side=tk.LEFT, padx=5)

        rf = ttk.LabelFrame(parent, text="Результаты (в реальном времени)", padding=8)
        rf.pack(fill=tk.BOTH, expand=True, padx=16, pady=6)
        self.results_text = scrolledtext.ScrolledText(
            rf, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9))
        self.results_text.pack(fill=tk.BOTH, expand=True)
        _bind_edit_menu(self.root, self.results_text, is_text=True, readonly=True)

        self.matched_count_var = tk.StringVar(value="Найдено уникальных пользователей: 0")
        ttk.Label(parent, textvariable=self.matched_count_var).pack(pady=3)

    def _build_mail_tab(self, parent):
        mf = ttk.LabelFrame(parent, text="Текст сообщения для рассылки", padding=10)
        mf.pack(fill=tk.X, padx=16, pady=10)
        self.mail_text = scrolledtext.ScrolledText(mf, height=7, wrap=tk.WORD)
        self.mail_text.pack(fill=tk.X)
        _bind_edit_menu(self.root, self.mail_text, is_text=True)

        of = ttk.Frame(parent); of.pack(fill=tk.X, padx=16, pady=5)
        ttk.Label(of, text="Задержка между сообщениями (сек):").pack(side=tk.LEFT)
        self.delay_var = tk.StringVar(value="3")
        delay_entry = ttk.Entry(of, textvariable=self.delay_var, width=6)
        delay_entry.pack(side=tk.LEFT, padx=6)
        _bind_edit_menu(self.root, delay_entry)

        self.send_btn = ttk.Button(
            parent,
            text="Отправить всем найденным пользователям",
            command=self._start_mailing,
            state=tk.DISABLED
        )
        self.send_btn.pack(pady=8)

        lf = ttk.LabelFrame(parent, text="Лог рассылки", padding=8)
        lf.pack(fill=tk.BOTH, expand=True, padx=16, pady=6)
        self.mail_log = scrolledtext.ScrolledText(
            lf, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9))
        self.mail_log.pack(fill=tk.BOTH, expand=True)
        _bind_edit_menu(self.root, self.mail_log, is_text=True, readonly=True)

        self.mail_progress_var = tk.StringVar()
        ttk.Label(parent, textvariable=self.mail_progress_var).pack(pady=3)

    # ── Async / thread-safe helpers ───────────────────────────────────────────

    def _run_async(self, coro):
        return asyncio.run_coroutine_threadsafe(coro, self.loop)

    def _ui(self, fn):
        """Queue a zero-arg callable to run on the Tk main thread."""
        self._ui_queue.append(fn)

    def _poll_ui_queue(self):
        while self._ui_queue:
            fn = self._ui_queue.pop(0)
            try:
                fn()
            except Exception:
                pass
        self.root.after(80, self._poll_ui_queue)

    def _set_status(self, text: str):
        self.status_var.set(text)

    # ── Config ────────────────────────────────────────────────────────────────

    def _load_config(self):
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, encoding="utf-8") as fh:
                cfg = json.load(fh)
            self.api_id_var.set(cfg.get("api_id", ""))
            self.api_hash_var.set(cfg.get("api_hash", ""))
            self.phone_var.set(cfg.get("phone", ""))
        except Exception:
            pass

    def _save_config(self):
        cfg = {
            "api_id":   self.api_id_var.get(),
            "api_hash": self.api_hash_var.get(),
            "phone":    self.phone_var.get(),
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, ensure_ascii=False, indent=2)

    # ── Connection ────────────────────────────────────────────────────────────

    def _connect(self):
        api_id_str = self.api_id_var.get().strip()
        api_hash   = self.api_hash_var.get().strip()
        phone      = self.phone_var.get().strip()
        if not all([api_id_str, api_hash, phone]):
            messagebox.showerror("Ошибка", "Заполните все поля!")
            return
        try:
            api_id = int(api_id_str)
        except ValueError:
            messagebox.showerror("Ошибка", "API ID должен быть числом!")
            return
        self._save_config()
        self._phone = phone
        self.connect_btn.config(state=tk.DISABLED)
        self._set_status("Подключение…")
        self._run_async(self._async_connect(api_id, api_hash, phone))

    async def _async_connect(self, api_id, api_hash, phone):
        from telethon import TelegramClient
        from telethon.errors import FloodWaitError
        try:
            self.client = TelegramClient(SESSION_FILE, api_id, api_hash)
            await self.client.connect()
            if not await self.client.is_user_authorized():
                try:
                    sent = await self.client.send_code_request(phone)
                    code_type = type(sent.type).__name__
                    if "App" in code_type:
                        hint = "Код отправлен в приложение Telegram — откройте Telegram на телефоне"
                    else:
                        hint = "Код отправлен по SMS"
                except FloodWaitError as e:
                    msg = f"Слишком много попыток. Подождите {e.seconds} сек. и попробуйте снова."
                    self._ui(lambda m=msg: self._set_status(m))
                    self._ui(lambda m=msg: messagebox.showwarning("FloodWait", m))
                    self._ui(lambda: self.connect_btn.config(state=tk.NORMAL))
                    return
                self._ui(lambda: self._code_frame.pack(fill=tk.X, padx=24, pady=8))
                self._ui(lambda h=hint: self._set_status(h))
            else:
                me = await self.client.get_me()
                self._ui(lambda m=me: self._on_connected(m))
        except Exception as ex:
            self._ui(lambda e=ex: self._set_status(f"Ошибка подключения: {e}"))
            self._ui(lambda: self.connect_btn.config(state=tk.NORMAL))

    def _resend_code(self):
        self._resend_btn.config(state=tk.DISABLED)
        self._set_status("Запрос нового кода…")
        self._run_async(self._async_resend_code())

    async def _async_resend_code(self):
        from telethon.errors import FloodWaitError
        try:
            sent = await self.client.send_code_request(self._phone)
            code_type = type(sent.type).__name__
            if "App" in code_type:
                hint = "Новый код отправлен в приложение Telegram"
            else:
                hint = "Новый код отправлен по SMS"
            self._ui(lambda h=hint: self._set_status(h))
        except FloodWaitError as e:
            msg = f"Слишком много попыток. Подождите {e.seconds} сек."
            self._ui(lambda m=msg: self._set_status(m))
            self._ui(lambda m=msg: messagebox.showwarning("FloodWait", m))
        except Exception as ex:
            self._ui(lambda e=ex: self._set_status(f"Ошибка: {e}"))
        finally:
            self._ui(lambda: self._resend_btn.config(state=tk.NORMAL))

    def _submit_code(self):
        code = self.code_var.get().strip()
        if not code:
            messagebox.showerror("Ошибка", "Введите код!")
            return
        self._run_async(self._async_submit_code(code))

    async def _async_submit_code(self, code):
        from telethon.errors import SessionPasswordNeededError
        try:
            await self.client.sign_in(self._phone, code)
            me = await self.client.get_me()
            self._ui(lambda: self._code_frame.pack_forget())
            self._ui(lambda m=me: self._on_connected(m))
        except SessionPasswordNeededError:
            self._ui(lambda: self._code_frame.pack_forget())
            self._ui(lambda: self._pw_frame.pack(fill=tk.X, padx=24, pady=8))
            self._ui(lambda: self._set_status("Требуется пароль 2FA"))
        except Exception as ex:
            self._ui(lambda e=ex: self._set_status(f"Ошибка кода: {e}"))

    def _submit_password(self):
        pw = self.password_var.get()
        self._run_async(self._async_submit_password(pw))

    async def _async_submit_password(self, pw):
        try:
            await self.client.sign_in(password=pw)
            me = await self.client.get_me()
            self._ui(lambda: self._pw_frame.pack_forget())
            self._ui(lambda m=me: self._on_connected(m))
        except Exception as ex:
            self._ui(lambda e=ex: self._set_status(f"Ошибка 2FA: {e}"))

    def _on_connected(self, me):
        name = f"{me.first_name or ''} {me.last_name or ''}".strip()
        tag  = f"@{me.username}" if me.username else me.phone
        self._set_status(f"Подключено: {name} ({tag})")
        self.connect_btn.config(state=tk.DISABLED)
        self.disconnect_btn.config(state=tk.NORMAL)
        self.start_btn.config(state=tk.NORMAL)
        self.send_btn.config(state=tk.NORMAL)

    def _disconnect(self):
        self._stop_monitoring()
        if self.client:
            self._run_async(self.client.disconnect())
        self.client = None
        self._set_status("Отключено")
        self.connect_btn.config(state=tk.NORMAL)
        self.disconnect_btn.config(state=tk.DISABLED)
        self.start_btn.config(state=tk.DISABLED)
        self.send_btn.config(state=tk.DISABLED)

    # ── Monitoring ────────────────────────────────────────────────────────────

    def _start_monitoring(self):
        kws = [" ".join(k.split()).lower() for k in self.keywords_var.get().split(",") if k.strip()]
        if not kws:
            messagebox.showerror("Ошибка", "Введите ключевые слова!")
            return
        self._keywords = kws
        self.is_monitoring = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self._set_status(f"Мониторинг активен  |  слова: {', '.join(kws)}")
        self._run_async(self._async_add_handler())

    async def _async_add_handler(self):
        from telethon import events

        async def on_message(event):
            if not self.is_monitoring:
                return
            raw  = event.message.text or event.message.message or ""
            text = " ".join(raw.split()).lower()   # нормализуем все виды пробелов
            if not any(kw in text for kw in self._keywords):
                return
            try:
                sender = await event.get_sender()
                chat   = await event.get_chat()
                s_name = _entity_name(sender)
                c_name = _chat_title(chat)

                if hasattr(sender, "id") and sender.id:
                    entry = {
                        "id":       sender.id,
                        "name":     s_name,
                        "username": getattr(sender, "username", None),
                        "chat":     c_name,
                    }
                    if not any(u["id"] == sender.id for u in self.matched_users):
                        self.matched_users.append(entry)

                ts       = datetime.now().strftime("%H:%M:%S")
                uname    = f"@{sender.username}" if getattr(sender, "username", None) else "нет username"
                line = f"[{ts}]  Чат: {c_name}  |  От: {s_name}  |  Username: {uname}  |  {text}\n"

                with open(RESULTS_FILE, "a", encoding="utf-8") as fh:
                    fh.write(line)

                self._ui(lambda l=line: self._add_result(l))
                self._ui(self._update_count)

            except Exception as ex:
                self._ui(lambda e=ex: self._add_result(f"[Ошибка обработки: {e}]\n"))

        self._active_handler = on_message
        self.client.add_event_handler(on_message, events.NewMessage)

    def _stop_monitoring(self):
        self.is_monitoring = False
        if self.client and self._active_handler:
            handler = self._active_handler
            self._run_async(self._async_remove_handler(handler))
        self._active_handler = None
        self.stop_btn.config(state=tk.DISABLED)
        self._save_summary()
        if self.client:
            self.start_btn.config(state=tk.NORMAL)
            self._set_status("Мониторинг остановлен")

    def _save_summary(self):
        if not self.matched_users:
            return
        ts_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sep    = "=" * 60
        lines  = [
            f"\n{sep}",
            f"СЕССИЯ ЗАВЕРШЕНА: {ts_now}",
            f"Найдено уникальных пользователей: {len(self.matched_users)}",
            sep,
        ]
        for i, u in enumerate(self.matched_users, 1):
            uname = f"@{u['username']}" if u["username"] else "нет username"
            lines.append(f"  {i}. {u['name']}  |  {uname}  |  Чат: {u['chat']}")
        lines.append(sep + "\n")
        with open(RESULTS_FILE, "a", encoding="utf-8") as fh:
            fh.write("\n".join(lines))

    async def _async_remove_handler(self, handler):
        self.client.remove_event_handler(handler)

    def _add_result(self, text: str):
        self.results_text.config(state=tk.NORMAL)
        self.results_text.insert(tk.END, text)
        self.results_text.see(tk.END)
        self.results_text.config(state=tk.DISABLED)

    def _update_count(self):
        self.matched_count_var.set(
            f"Найдено уникальных пользователей: {len(self.matched_users)}"
        )

    def _clear_results(self):
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.config(state=tk.DISABLED)
        self.matched_users.clear()
        self._update_count()

    # ── Mailing ───────────────────────────────────────────────────────────────

    def _start_mailing(self):
        if not self.matched_users:
            messagebox.showinfo("Рассылка",
                "Список пуст. Сначала запустите мониторинг и дождитесь совпадений.")
            return
        msg = self.mail_text.get(1.0, tk.END).strip()
        if not msg:
            messagebox.showerror("Ошибка", "Введите текст сообщения!")
            return
        try:
            delay = max(0.5, float(self.delay_var.get()))
        except ValueError:
            delay = 3.0
        preview = msg[:80] + ("…" if len(msg) > 80 else "")
        if not messagebox.askyesno(
            "Подтверждение",
            f"Отправить сообщение {len(self.matched_users)} пользователям?\n\n«{preview}»"
        ):
            return
        self.send_btn.config(state=tk.DISABLED)
        self._run_async(self._async_mailing(msg, delay))

    async def _async_mailing(self, message: str, delay: float):
        import asyncio as aio
        total   = len(self.matched_users)
        success = 0
        for i, user in enumerate(self.matched_users, 1):
            try:
                await self.client.send_message(user["id"], message)
                log = f"[{i}/{total}] ✓  {user['name']}\n"
                success += 1
            except Exception as ex:
                log = f"[{i}/{total}] ✗  {user['name']}: {ex}\n"
            self._ui(lambda l=log: self._add_mail_log(l))
            self._ui(lambda v=f"Прогресс: {i}/{total}": self.mail_progress_var.set(v))
            if i < total:
                await aio.sleep(delay)
        summary = f"\n--- Рассылка завершена. Успешно: {success}/{total} ---\n"
        self._ui(lambda s=summary: self._add_mail_log(s))
        self._ui(lambda: self.send_btn.config(state=tk.NORMAL))

    def _add_mail_log(self, text: str):
        self.mail_log.config(state=tk.NORMAL)
        self.mail_log.insert(tk.END, text)
        self.mail_log.see(tk.END)
        self.mail_log.config(state=tk.DISABLED)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
