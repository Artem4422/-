const tg = window.Telegram?.WebApp;
if (tg) {
  tg.ready();
  tg.expand();
}

const API = "";
const headers = () => {
  const h = { "Content-Type": "application/json" };
  if (tg?.initData) h["X-Telegram-Init-Data"] = tg.initData;
  return h;
};

async function api(path, opts = {}) {
  const res = await fetch(API + path, {
    ...opts,
    headers: { ...headers(), ...(opts.headers || {}) },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.detail || data.message || res.statusText);
  return data;
}

function toast(msg) {
  if (tg?.showAlert) tg.showAlert(msg);
  else alert(msg);
}

function switchTab(name) {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
  document.querySelectorAll(".panel").forEach((p) => p.classList.toggle("active", p.id === `tab-${name}`));
}

document.querySelectorAll(".tab").forEach((btn) => {
  btn.onclick = () => switchTab(btn.dataset.tab);
});

let state = {};
let leads = [];

function renderState(s) {
  state = s;
  document.getElementById("statusLine").textContent = s.status || "—";
  document.getElementById("connInfo").textContent = s.connected_user
    ? `✅ ${s.connected_user.name} (${s.connected_user.tag})`
    : s.auth_hint || "Не подключено";

  const authBlock = document.getElementById("authBlock");
  authBlock.classList.toggle("hidden", !s.auth_state);
  document.getElementById("scanProgress").textContent = s.scan_progress || "";
  document.getElementById("mailProgress").textContent = s.mail_progress || "";
  document.getElementById("keywords").value = s.keywords || "";
  document.getElementById("scanHistory").checked = !!s.scan_history;
  document.getElementById("historyLimit").value = s.history_limit || "0";
  document.getElementById("mailMessage").value = s.mail_message || "";
  document.getElementById("mailDelay").value = s.mail_delay || "3";
  document.getElementById("resultsLog").textContent = (s.results_tail || []).join("\n") || "—";
  document.getElementById("mailLog").textContent = (s.mail_log_tail || []).join("\n") || "—";
  document.getElementById("leadCount").textContent = `(${s.selected_count}/${s.leads_count})`;
}

function renderLeads(list) {
  leads = list;
  const box = document.getElementById("leadsList");
  box.innerHTML = "";
  list.slice(0, 200).forEach((u) => {
    const el = document.createElement("div");
    el.className = "lead" + (u.selected ? " on" : "");
    el.innerHTML = `
      <div class="box"></div>
      <div>
        <div><strong>${esc(u.name)}</strong> ${u.username ? "@" + esc(u.username) : ""}</div>
        <div class="meta">${esc(u.chat || "")} · ${esc(u.source || "")}</div>
        <div class="meta">${esc((u.message || "").slice(0, 100))}</div>
      </div>`;
    el.onclick = async () => {
      await api("/api/leads/toggle", { method: "POST", body: JSON.stringify({ key: u._key }) });
      await refresh();
    };
    box.appendChild(el);
  });
  if (list.length > 200) {
    const more = document.createElement("p");
    more.className = "meta";
    more.textContent = `… показано 200 из ${list.length}`;
    box.appendChild(more);
  }
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

async function refresh() {
  const s = await api("/api/status");
  renderState(s);
  const l = await api("/api/leads");
  renderLeads(l);
}

async function loadConfig() {
  const cfg = await api("/api/config");
  document.getElementById("apiId").value = cfg.api_id || "";
  document.getElementById("apiHash").value = cfg.api_hash || "";
  document.getElementById("phone").value = cfg.phone || "";
}

document.getElementById("saveConfig").onclick = async () => {
  try {
    await api("/api/config", {
      method: "POST",
      body: JSON.stringify({
        api_id: document.getElementById("apiId").value.trim(),
        api_hash: document.getElementById("apiHash").value.trim(),
        phone: document.getElementById("phone").value.trim(),
      }),
    });
    toast("API сохранён");
    await refresh();
  } catch (e) { toast(e.message); }
};

document.getElementById("resetSession").onclick = async () => {
  if (!confirm("Сбросить сессию Telegram?")) return;
  await api("/api/reset/session", { method: "POST" });
  await refresh();
};

document.getElementById("resetAll").onclick = async () => {
  if (!confirm("Сбросить API и сессию?")) return;
  await api("/api/reset/all", { method: "POST" });
  await loadConfig();
  await refresh();
};

document.getElementById("connectBtn").onclick = async () => {
  try { await api("/api/connect", { method: "POST" }); await refresh(); }
  catch (e) { toast(e.message); }
};
document.getElementById("disconnectBtn").onclick = async () => {
  await api("/api/disconnect", { method: "POST" });
  await refresh();
};
document.getElementById("submitCode").onclick = async () => {
  try {
    await api("/api/auth/code", { method: "POST", body: JSON.stringify({ code: document.getElementById("authCode").value }) });
    await refresh();
  } catch (e) { toast(e.message); }
};
document.getElementById("submit2fa").onclick = async () => {
  try {
    await api("/api/auth/2fa", { method: "POST", body: JSON.stringify({ password: document.getElementById("auth2fa").value }) });
    await refresh();
  } catch (e) { toast(e.message); }
};
document.getElementById("resendCode").onclick = async () => {
  try { await api("/api/auth/resend", { method: "POST" }); await refresh(); }
  catch (e) { toast(e.message); }
};

document.getElementById("saveMonSettings").onclick = async () => {
  await api("/api/settings", {
    method: "POST",
    body: JSON.stringify({
      keywords: document.getElementById("keywords").value,
      scan_history: document.getElementById("scanHistory").checked,
      history_limit: document.getElementById("historyLimit").value,
    }),
  });
  toast("Сохранено");
  await refresh();
};
document.getElementById("startMon").onclick = async () => {
  await api("/api/settings", {
    method: "POST",
    body: JSON.stringify({ keywords: document.getElementById("keywords").value }),
  });
  try { await api("/api/monitor/start", { method: "POST" }); await refresh(); }
  catch (e) { toast(e.message); }
};
document.getElementById("scanHist").onclick = async () => {
  try { await api("/api/monitor/scan", { method: "POST" }); await refresh(); }
  catch (e) { toast(e.message); }
};
document.getElementById("stopMon").onclick = async () => {
  await api("/api/monitor/stop", { method: "POST" });
  await refresh();
};

document.getElementById("addManual").onclick = async () => {
  try {
    await api("/api/leads/manual", {
      method: "POST",
      body: JSON.stringify({ text: document.getElementById("manualUsers").value }),
    });
    document.getElementById("manualUsers").value = "";
    await refresh();
  } catch (e) { toast(e.message); }
};

document.getElementById("xlsxFile").onchange = async (ev) => {
  const file = ev.target.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);
  const h = {};
  if (tg?.initData) h["X-Telegram-Init-Data"] = tg.initData;
  const res = await fetch("/api/leads/import", { method: "POST", headers: h, body: fd });
  const data = await res.json();
  if (!res.ok) return toast(data.detail || "Ошибка импорта");
  toast(`Импорт: ${data.found}, добавлено ${data.added}`);
  ev.target.value = "";
  await refresh();
};

document.getElementById("removeImported").onclick = async () => {
  await api("/api/leads/imported", { method: "DELETE" });
  await refresh();
};
document.getElementById("clearLeads").onclick = async () => {
  if (!confirm("Очистить всех лидов?")) return;
  await api("/api/leads", { method: "DELETE" });
  await refresh();
};
document.getElementById("selectAll").onclick = async () => {
  await api("/api/leads/selection", { method: "POST", body: JSON.stringify({ keys: null, selected: true }) });
  await refresh();
};
document.getElementById("selectNone").onclick = async () => {
  await api("/api/leads/selection", { method: "POST", body: JSON.stringify({ keys: null, selected: false }) });
  await refresh();
};

document.getElementById("sendMail").onclick = async () => {
  if (!confirm("Отправить рассылку выбранным?")) return;
  try {
    const r = await api("/api/mail/send", {
      method: "POST",
      body: JSON.stringify({
        message: document.getElementById("mailMessage").value,
        delay: parseFloat(document.getElementById("mailDelay").value),
      }),
    });
    toast(`Готово: ${r.success}/${r.total}`);
    await refresh();
  } catch (e) { toast(e.message); }
};

loadConfig().then(refresh).catch((e) => toast(e.message));
setInterval(refresh, 5000);
