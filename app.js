/* =====================================================================
   Weekly Shop — meal planner + aisle-sorted shopping list
   Ported from the Food_Improved_v4.xlsm macro (RefreshShoppingList).
   ===================================================================== */

const MAX_MEALS = 8;               // change here if you ever want more meal slots
const LS_STATE = "weeklyshop.state.v1";
const LS_SETTINGS = "weeklyshop.settings.v1";
const LS_SYNCED = "weeklyshop.syncedAt.v1";

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

/* ---------------------------------------------------------------------
   The macro's algorithm (pure functions — no screen access)
   --------------------------------------------------------------------- */

// Macro: LCase(Replace(Replace(item,"(have)",""),"?","")) then Trim
function cleanItem(s) {
  return String(s || "").replace(/\(have\)/gi, "").replace(/\?/g, "").replace(/\s+/g, " ").trim().toLowerCase();
}

// Macro: FindAisle — keyword is "contained in" the ingredient. The macro took the
// first match down the sheet; here the LONGEST matching keyword wins (ties: earlier
// one), so a specific keyword you add later (e.g. "green beans") beats a generic
// one that was already there (e.g. "bean").
function findAisle(name, st) {
  let best = null, len = 0;
  for (const kw of st.keywords) {
    if (kw.k && name.includes(kw.k) && kw.k.length > len) { best = kw.a; len = kw.k.length; }
  }
  return best;
}

function aisleRank(aisle, st) {
  const i = aisle == null ? -1 : st.aisles.indexOf(aisle);
  return i < 0 ? 999 : i;
}

function collectIngredients(st) {
  const out = [];
  st.meals.filter((m) => !m.reuse).forEach((m) => String(m.ingredients || "").split(/\r?\n/).forEach((l) => out.push(l)));
  st.misc.filter((x) => x.selected).forEach((x) => out.push(x.name));   // misc only if "yellow"
  return out;
}

function buildShopping(st) {
  const counts = new Map();
  for (const raw of collectIngredients(st)) {
    const n = cleanItem(raw);
    if (!n) continue;
    counts.set(n, (counts.get(n) || 0) + 1);
  }
  const items = [...counts].map(([name, count]) => ({ name, count, aisle: findAisle(name, st), checked: false }));
  items.sort((a, b) => {
    const d = aisleRank(a.aisle, st) - aisleRank(b.aisle, st);
    if (d !== 0) return d;
    return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
  });
  return items;
}

/* ---------------------------------------------------------------------
   State
   --------------------------------------------------------------------- */
/* ---------------------------------------------------------------------
   Previous meals (history) — pure functions
   --------------------------------------------------------------------- */
const mealKey = (n) => String(n || "").replace(/\s+/g, " ").trim().toLowerCase();
// Ingredient text as saved in history: no "(have)" / "?" markers (those are one-week notes).
const ingDisplay = (s) => String(s || "").replace(/\(have\)/gi, "").replace(/\?/g, "").replace(/\s+/g, " ").trim();

// Week = Monday..Sunday; a meal counts once per week however many times you populate.
function weekKey(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  x.setDate(x.getDate() - ((x.getDay() + 6) % 7));
  return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`;
}

function findHistory(st, name) {
  const k = mealKey(name); if (!k) return null;
  return st.history.find((h) => mealKey(h.name) === k) || null;
}

// Save this week's meals into history: add new meals, ADD any new ingredients to existing
// meals (never duplicate, never remove), bump "times chosen" once per week.
function recordMeals(st, now = new Date()) {
  const wk = weekKey(now), seen = new Set();
  let meals = 0, newMeals = 0, added = 0;
  for (const m of st.meals) {
    if (m.reuse) continue;                 // carried over from last week: already recorded then
    const key = mealKey(m.name);
    if (!key || seen.has(key)) continue;
    seen.add(key); meals++;
    let h = findHistory(st, m.name);
    if (!h) {
      h = { id: "h" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name: String(m.name).replace(/\s+/g, " ").trim(), book: "", page: "", ingredients: [], count: 0, lastUsed: null, lastCountedWeek: null };
      st.history.push(h); newMeals++;
    }
    if (m.book && m.book.trim()) h.book = m.book.trim();
    if (m.page && String(m.page).trim()) h.page = String(m.page).trim();
    const have = new Set(h.ingredients.map(cleanItem));
    String(m.ingredients || "").split(/\r?\n/).forEach((line) => {
      const k = cleanItem(line);
      if (k && !have.has(k)) { have.add(k); h.ingredients.push(ingDisplay(line)); added++; }
    });
    if (h.lastCountedWeek !== wk) { h.count++; h.lastCountedWeek = wk; }
    h.lastUsed = now.toISOString();
  }
  return { meals, newMeals, added };
}

function top5(st) {
  return st.history.filter((h) => h.count > 0).slice()
    .sort((a, b) => b.count - a.count || String(b.lastUsed).localeCompare(String(a.lastUsed)) || a.name.localeCompare(b.name)).slice(0, 5);
}

let state;
const ui = { screen: "home", openMeals: new Set([0]), editMisc: false, shopBanner: "", kwFilter: "", kwOpen: false, settingsMsg: null, openHist: new Set(), histFilter: "", auto: {}, autoNote: {} };

function emptyMeal() { return { name: "", book: "", page: "", ingredients: "", reuse: false }; }

function freshState() {
  const s = JSON.parse(JSON.stringify(SEED));
  s.version = 1; s.updatedAt = null; s.shopping = null;
  return s;
}

function normalise() {
  state.usedMeals = Array.isArray(state.usedMeals) ? state.usedMeals : [];
  state.aisles = state.aisles || [];
  state.keywords = state.keywords || [];
  state.misc = state.misc || [];
  state.history = Array.isArray(state.history) ? state.history.map((h) => ({
    id: h.id || "h" + Math.random().toString(36).slice(2, 9), name: h.name || "", book: h.book || "", page: h.page || "",
    ingredients: Array.isArray(h.ingredients) ? h.ingredients : [], count: h.count || 0, lastUsed: h.lastUsed || null, lastCountedWeek: h.lastCountedWeek || null
  })) : [];
  state.meals = (state.meals || []).map((m) => ({ ...emptyMeal(), ...m }));
  while (state.meals.length < MAX_MEALS) state.meals.push(emptyMeal());
  state.meals = state.meals.map((m) => {       // older data: a meal flagged "used" moves out of the planner
    if (m.used) { state.usedMeals.push({ name: m.name, book: m.book, page: m.page, ingredients: m.ingredients }); return emptyMeal(); }
    const { used, ...rest } = m; return rest;
  });
  if (state.shopping && !Array.isArray(state.shopping.items)) state.shopping = null;
}

function loadState() {
  try {
    const raw = localStorage.getItem(LS_STATE);
    if (raw) { state = JSON.parse(raw); normalise(); return; }
  } catch (e) {}
  state = freshState(); normalise();
}

function save() { try { localStorage.setItem(LS_STATE, JSON.stringify(state)); } catch (e) {} }

// Any real change: stamp it, store it, and queue an automatic backup.
function touch() { state.updatedAt = new Date().toISOString(); save(); scheduleBackup(); renderSyncLines(); }

/* ---------------------------------------------------------------------
   Settings (GitHub backup details live here, separate from the data)
   --------------------------------------------------------------------- */
function defaultSettings() { return { github: { token: "", owner: "kjrmthdj6w-hue", repo: "weekly-shop-data", path: "meal-data.json", auto: true } }; }
function getSettings() {
  const d = defaultSettings();
  try { const p = JSON.parse(localStorage.getItem(LS_SETTINGS)); if (p && p.github) { const def = d.github; d.github = { ...def, ...p.github }; if (!d.github.owner) d.github.owner = def.owner; if (!d.github.repo) d.github.repo = def.repo; } } catch (e) {}
  return d;
}
function putSettings(s) { try { localStorage.setItem(LS_SETTINGS, JSON.stringify(s)); } catch (e) {} }
const getSynced = () => { try { return localStorage.getItem(LS_SYNCED) || null; } catch (e) { return null; } };
const setSynced = (v) => { try { v ? localStorage.setItem(LS_SYNCED, v) : localStorage.removeItem(LS_SYNCED); } catch (e) {} };

/* ---------------------------------------------------------------------
   GitHub backup / restore
   --------------------------------------------------------------------- */
let syncInfo = { text: "", kind: "" };
let backupTimer = null;

function ghCfg() { const g = getSettings().github; return g.token && g.owner && g.repo ? g : null; }
function ghUrl(g) {
  const path = (g.path || "meal-data.json").split("/").map(encodeURIComponent).join("/");
  return `https://api.github.com/repos/${encodeURIComponent(g.owner)}/${encodeURIComponent(g.repo)}/contents/${path}`;
}
const ghHeaders = (g) => ({ Authorization: `Bearer ${g.token}`, Accept: "application/vnd.github+json" });

async function timedFetch(url, opts = {}, ms = 9000) {
  const c = new AbortController(); const t = setTimeout(() => c.abort(), ms);
  try { return await fetch(url, { ...opts, signal: c.signal }); }
  catch (e) { throw new Error(e.name === "AbortError" ? "GitHub took too long to answer" : "couldn't reach GitHub (offline?)"); }
  finally { clearTimeout(t); }
}

async function ghGet(g) {
  const res = await timedFetch(ghUrl(g), { headers: ghHeaders(g), cache: "no-store" });
  if (res.status === 404) return null;
  if (res.status === 401) throw new Error("token rejected — check it in Settings");
  if (!res.ok) throw new Error(`GitHub error ${res.status} — check username/repo/token`);
  const j = await res.json();
  const text = decodeURIComponent(escape(atob(String(j.content).replace(/\n/g, ""))));
  return { sha: j.sha, data: JSON.parse(text) };
}

async function ghPush(force = false) {
  const g = ghCfg(); if (!g) throw new Error("add your GitHub details in Settings first");
  if (!state.updatedAt) { state.updatedAt = new Date().toISOString(); save(); }
  const remote = await ghGet(g);
  if (remote && !force) {
    const rAt = remote.data && remote.data.updatedAt;
    if (rAt && rAt !== getSynced() && rAt !== state.updatedAt) { const e = new Error("conflict"); e.conflict = true; throw e; }
  }
  const body = {
    message: `Weekly Shop backup — ${new Date().toISOString()}`,
    content: btoa(unescape(encodeURIComponent(JSON.stringify(state, null, 1)))),
    ...(remote ? { sha: remote.sha } : {})
  };
  const res = await timedFetch(ghUrl(g), { method: "PUT", headers: { ...ghHeaders(g), "Content-Type": "application/json" }, body: JSON.stringify(body) });
  if (!res.ok) { const j = await res.json().catch(() => ({})); throw new Error(j.message || `GitHub error ${res.status}`); }
  setSynced(state.updatedAt);
}

function setSync(text, kind) { syncInfo = { text, kind }; renderSyncLines(); }
function renderSyncLines() {
  const base = ghCfg() ? (state.updatedAt && state.updatedAt !== getSynced() ? "Changes not backed up yet" : "All changes backed up to GitHub") : "Backup off — set up in Settings";
  const t = syncInfo.text || base;
  document.querySelectorAll(".syncline").forEach((el) => { el.textContent = t; });
}

// Backup now; asks before overwriting a newer copy from somewhere else.
async function backupNow({ quiet = false } = {}) {
  setSync("Backing up…", "");
  try {
    await ghPush(false);
    setSync(`✅ Backed up ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`, "ok");
    return { ok: true };
  } catch (e) {
    if (e.conflict) {
      if (quiet) { setSync("⚠️ GitHub has a different newer copy — open Settings to choose", "err"); return { ok: false, conflict: true }; }
      if (confirm("GitHub holds a newer copy (from another device or before a reset).\n\nOK = overwrite it with THIS phone's data.\nCancel = keep GitHub's copy (use Settings → Load saved data).")) {
        try { await ghPush(true); setSync("✅ Backed up", "ok"); return { ok: true }; } catch (e2) { setSync("Backup failed: " + e2.message, "err"); return { ok: false, error: e2.message }; }
      }
      setSync("Backup skipped — GitHub copy kept", "err"); return { ok: false, conflict: true };
    }
    setSync("Backup failed: " + e.message, "err");
    return { ok: false, error: e.message };
  }
}

// Back up straight away (instead of waiting 10s) — used for deliberate actions like marking a meal used.
function backupSoon() { const g = ghCfg(); if (!g || !g.auto) return; clearTimeout(backupTimer); backupNow({ quiet: true }); }

function scheduleBackup() {
  const g = ghCfg(); if (!g || !g.auto) return;
  clearTimeout(backupTimer);
  backupTimer = setTimeout(() => backupNow({ quiet: true }), 10000);   // 10s after the last change
}

// On opening: if GitHub has something newer and this phone has no unsaved changes, take it.
async function syncOnOpen() {
  const g = ghCfg(); if (!g) return;
  try {
    const r = await ghGet(g);
    if (!r || !r.data) return;
    const rAt = r.data.updatedAt;
    if (!rAt || rAt === getSynced()) return;
    if (state.updatedAt === getSynced()) {
      adopt(r.data);
      toast("Loaded your latest saved data ✓");
    } else {
      setSync("⚠️ GitHub has a different copy — Settings → Load saved data to take it", "err");
    }
  } catch (e) { /* offline is fine */ }
}

// Home-screen "Update" button: pull the newest backup (e.g. after using another device).
async function updateFromBackup() {
  const g = ghCfg(), msg = $("upd-msg"), say = (t, k) => { if (msg) { msg.textContent = t; msg.style.color = k === "err" ? "var(--danger)" : k === "ok" ? "var(--accent)" : ""; } };
  if (!g) return say("Backup isn't set up yet — add your GitHub details in Settings.", "err");
  say("Checking GitHub…", "");
  try {
    const r = await ghGet(g);
    if (!r || !r.data) return say("No backup found on GitHub yet.", "err");
    const dirty = state.updatedAt && state.updatedAt !== getSynced();
    if (r.data.updatedAt === state.updatedAt || (!dirty && r.data.updatedAt === getSynced())) return say("✅ Already up to date", "ok");
    if (dirty && !confirm("This device has changes that aren't backed up yet. Updating will replace them with the latest backup. Continue?")) return say("Update cancelled — your changes are kept.", "");
    adopt(r.data);
    const m2 = $("upd-msg"); if (m2) { m2.textContent = "✅ Updated from the latest backup"; m2.style.color = "var(--accent)"; }
  } catch (e) { say("Couldn't reach GitHub: " + (e.message || "offline"), "err"); }
}

function adopt(data) {
  state = data; normalise(); save(); setSynced(state.updatedAt || null);
  syncInfo = { text: "", kind: "" }; render();
}

async function loadSaved() {
  const msg = $("load-msg"); const g = ghCfg();
  const show = (cls, t) => { msg.className = cls; msg.textContent = t; };
  if (!g) return show("err", "Add your token, username and repo first (and tap Save).");
  show("hint", "Loading saved data…");
  try {
    const r = await ghGet(g);
    if (!r) return show("err", "No backup found in that repo yet — tap Backup now first.");
    if (state.updatedAt && state.updatedAt !== getSynced() && !confirm("This phone has changes that aren't backed up. Loading will replace them with the GitHub copy. Continue?")) return show("hint", "Cancelled — nothing changed.");
    adopt(r.data);
    const n = state.meals.filter((m) => m.name.trim()).length;
    ui.settingsMsg = `✅ Data loaded successfully — ${n} planned meal${n === 1 ? "" : "s"}, ${state.keywords.length} keywords, ${state.shopping ? state.shopping.items.length : 0} shopping items`;
    go("settings");
  } catch (e) { show("err", "Failed to load: " + e.message); }
}

/* ---------------------------------------------------------------------
   Navigation
   --------------------------------------------------------------------- */
const SCREENS = ["home", "shopping", "planner", "aisles", "recipes", "history", "settings"];
function go(name) {
  ui.screen = name;
  SCREENS.forEach((s) => $("s-" + s).classList.toggle("active", s === name));
  render();
  window.scrollTo(0, 0);
}
function render() {
  ({ home: rHome, shopping: rShopping, planner: rPlanner, aisles: rAisles, recipes: rRecipes, history: rHistory, settings: rSettings })[ui.screen]();
  renderSyncLines();
}
const bar = (title) => `<div class="bar"><button class="back" onclick="go('home')">‹ Home</button><h1>${esc(title)}</h1></div>`;
let toastTimer;
function toast(msg) { const t = $("toast"); t.textContent = msg; t.classList.add("show"); clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove("show"), 2600); }
const aisleOptions = (sel) => state.aisles.map((a) => `<option ${a === sel ? "selected" : ""}>${esc(a)}</option>`).join("");

/* ---------------------------------------------------------------------
   HOME
   --------------------------------------------------------------------- */
function rHome() {
  const sh = state.shopping, n = sh ? sh.items.length : 0, ticked = sh ? sh.items.filter((i) => i.checked).length : 0;
  const planned = state.meals.filter((m) => m.name.trim()).length;
  $("v-home").innerHTML = `
    <div class="brand"><h1>Weekly<br><span>Shop</span></h1><button class="gear" aria-label="Settings" onclick="go('settings')">⚙︎</button></div>
    <button class="btn ghost updbtn" id="upd-btn" onclick="updateFromBackup()">⟳ Update from latest backup</button>
    <div class="syncline" id="upd-msg" style="margin:-4px 0 14px"></div>
    <div class="tiles">
      <button class="tile main" onclick="go('shopping')"><span class="ic">🛒</span><span><div class="tt">Shopping list</div><div class="ts">${sh ? `${n} items · ${ticked} ticked` : "Not populated yet"}</div></span></button>
      <button class="tile" onclick="go('planner')"><span class="ic">🍽️</span><span><div class="tt">Meal planner</div><div class="ts">${planned} of ${MAX_MEALS} meals planned</div></span></button>
      <button class="tile" onclick="go('aisles')"><span class="ic">🧭</span><span><div class="tt">Aisle planner</div><div class="ts">${state.aisles.length} aisles · ${state.keywords.length} keywords</div></span></button>
      <button class="tile" onclick="go('recipes')"><span class="ic">📖</span><span><div class="tt">Weekly recipes</div><div class="ts">${planned} recipe${planned === 1 ? "" : "s"} this week</div></span></button>
      <button class="tile wide" onclick="go('history')"><span class="ic">🕘</span><span><div class="tt">Previous meals</div><div class="ts">${state.history.length ? `${state.history.length} meal${state.history.length === 1 ? "" : "s"} remembered · top: ${esc(top5(state)[0] ? top5(state)[0].name : "–")}` : "Builds up as you plan"}</div></span></button>
    </div>
    <div class="syncline"></div>`;
}

/* ---------------------------------------------------------------------
   SHOPPING LIST
   --------------------------------------------------------------------- */
function shopGroups() {
  const items = state.shopping.items, g = [];
  state.aisles.forEach((a) => {
    const idx = items.map((it, i) => (it.aisle === a ? i : -1)).filter((i) => i >= 0);
    if (idx.length) g.push({ aisle: a, idx });
  });
  const other = items.map((it, i) => (!it.aisle || !state.aisles.includes(it.aisle) ? i : -1)).filter((i) => i >= 0);
  if (other.length) g.push({ aisle: null, idx: other });
  return g;
}

function rShopping() {
  const el = $("v-shopping");
  if (!state.shopping || !state.shopping.items.length) {
    el.innerHTML = bar("Shopping list") + (ui.shopBanner ? `<div class="ok">${esc(ui.shopBanner)}</div>` : "") +
      `<div class="empty"><div class="big">🛒</div><b>Nothing on the list yet</b><p>Plan your meals, then tap <b>Complete – populate shopping list</b>.</p><button class="btn" onclick="go('planner')">Go to Meal planner</button></div>`;
    ui.shopBanner = ""; return;
  }
  const items = state.shopping.items, unknown = items.filter((i) => !i.aisle || !state.aisles.includes(i.aisle)).length;
  const when = new Date(state.shopping.generatedAt || Date.now()).toLocaleString([], { weekday: "short", day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" });
  const ticked = items.filter((i) => i.checked).length;
  const groups = shopGroups().map((g) => `
    <div class="aisle ${g.aisle ? "" : "unknown"}"><h3><span>${g.aisle ? esc(g.aisle) : "⚠ Needs an aisle"}</span><span>${g.idx.length}</span></h3>
    <div class="list">${g.idx.map((i) => {
      const it = items[i];
      return it.aisle && state.aisles.includes(it.aisle)
        ? `<label class="item ${it.checked ? "done" : ""}"><input type="checkbox" ${it.checked ? "checked" : ""} onchange="tick(${i},this)"><span class="nm">${esc(it.name)}</span>${it.count > 1 ? `<span class="qty">×${it.count}</span>` : ""}</label>`
        : `<div class="item unk"><span class="nm">${esc(it.name)}${it.count > 1 ? ` ×${it.count}` : ""}</span><select onchange="assignAisle(${i},this.value)"><option value="">Choose aisle…</option>${aisleOptions("")}</select></div>`;
    }).join("")}</div></div>`).join("");
  el.innerHTML = bar("Shopping list") +
    (ui.shopBanner ? `<div class="ok">${esc(ui.shopBanner)}</div>` : "") +
    (unknown ? `<div class="warn">${unknown} item${unknown === 1 ? " needs" : "s need"} an aisle — pick one below and it's remembered for next time.</div>` : "") +
    `<p class="sub" id="shop-sub">Made ${esc(when)} · ${items.length} items · <span id="tick-n">${ticked}</span> ticked</p>` + groups +
    `<div class="stack"><button class="btn ghost" onclick="shareList()">Share list</button><button class="btn ghost" onclick="untickAll()">Untick everything</button></div>
     <div class="syncline"></div>`;
  ui.shopBanner = "";
}

function tick(i, box) {
  state.shopping.items[i].checked = box.checked; save();          // local only — not worth a backup each tap
  box.closest(".item").classList.toggle("done", box.checked);
  const n = $("tick-n"); if (n) n.textContent = state.shopping.items.filter((x) => x.checked).length;
}
function untickAll() { state.shopping.items.forEach((i) => (i.checked = false)); save(); rShopping(); }

function assignAisle(i, aisle) {
  if (!aisle) return;
  addKeyword(state.shopping.items[i].name, aisle);
  reclassify(); touch(); render(); toast(`Remembered — "${state.shopping.items[i].name}" → ${aisle}`);
}

function shareList() {
  const lines = [];
  shopGroups().forEach((g) => {
    lines.push((g.aisle || "Needs an aisle").toUpperCase());
    g.idx.forEach((i) => { const it = state.shopping.items[i]; lines.push(`${it.checked ? "☑" : "☐"} ${it.name}${it.count > 1 ? " x" + it.count : ""}`); });
    lines.push("");
  });
  const text = lines.join("\n").trim();
  if (navigator.share) navigator.share({ title: "Shopping list", text }).catch(() => {});
  else { prompt("Copy your list:", text); }
}

/* ---------------------------------------------------------------------
   MEAL PLANNER
   --------------------------------------------------------------------- */
function mealSummary(m) {
  const n = String(m.ingredients || "").split(/\r?\n/).filter((l) => l.trim()).length;
  const bits = [m.book && m.book.trim(), m.page && m.page.toString().trim() ? "p. " + m.page.toString().trim() : ""].filter(Boolean).join(", ");
  return [m.reuse ? "♻ Re-used" : "", bits, n ? `${n} ingredient${n === 1 ? "" : "s"}` : ""].filter(Boolean).join(" · ") || "Tap to add";
}

function rPlanner() {
  const books = [...new Set(state.meals.map((m) => (m.book || "").trim()).filter(Boolean))];
  const meals = state.meals.slice(0, Math.max(MAX_MEALS, state.meals.length)).map((m, i) => `
    <div class="meal ${m.name.trim() ? "filled" : ""} ${m.reuse ? "reuse" : ""} ${ui.openMeals.has(i) ? "open" : ""}" id="meal-${i}">
      <button class="mhead" onclick="toggleMeal(${i})"><span class="mnum">${i + 1}</span>
        <span class="mtxt"><span class="mt" id="mt-${i}">${esc(m.name.trim() || "Meal " + (i + 1))}</span><span class="ms" id="ms-${i}">${esc(mealSummary(m))}</span></span><span class="chev">›</span></button>
      <div class="mbody">
        <label class="l">Meal name</label><input type="text" id="mn-${i}" value="${esc(m.name)}" autocomplete="off" oninput="onNameInput(${i},this.value)" onchange="onNameCommit(${i})" onfocus="renderSugs(${i})" placeholder="e.g. Greek salad">
        <div class="sugs" id="sug-${i}"></div><div class="hint" id="auto-${i}">${ui.autoNote[i] ? AUTO_NOTE : ""}</div>
        <div class="row2"><div><label class="l">Recipe book</label><input type="text" id="mb-${i}" list="books" value="${esc(m.book)}" oninput="editMeal(${i},'book',this.value)" placeholder="e.g. Fast food"></div>
        <div class="narrow"><label class="l">Page</label><input type="text" id="mp-${i}" inputmode="numeric" value="${esc(m.page)}" oninput="editMeal(${i},'page',this.value)" placeholder="18"></div></div>
        <label class="switch reusebox"><span><b>Meal to be re-used this week</b><small>Not cooked last week — ingredients are already bought, so they're left off the shopping list.</small></span><input type="checkbox" id="mr-${i}" ${m.reuse ? "checked" : ""} onchange="setReuse(${i},this.checked)"></label>
        <label class="l">Ingredients <span class="hintinline" id="icn-${i}"></span></label>
        <div class="chips ingchips" id="ic-${i}"></div>
        <textarea id="mi-${i}" oninput="editMeal(${i},'ingredients',this.value)" placeholder="Add ingredients here, one per line">${esc(m.ingredients)}</textarea>
        <button class="link" onclick="clearMeal(${i})">Clear this meal</button>
      </div></div>`).join("");
  const misc = state.misc.map((x, i) => `<button class="chip ${x.selected ? "on" : ""}" onclick="${ui.editMisc ? `delMisc(${i})` : `toggleMisc(${i})`}">${esc(x.name)}${ui.editMisc ? '<span class="x">✕</span>' : ""}</button>`).join("");
  $("v-planner").innerHTML = bar("Meal planner") +
    `<p class="sub">Fill in up to ${MAX_MEALS} meals, tick the extras you need, then populate the list.</p>
     <datalist id="books">${books.map((b) => `<option value="${esc(b)}">`).join("")}</datalist>` + meals +
    `<div class="secthead"><h2>Extras &amp; household</h2><button class="link" onclick="ui.editMisc=!ui.editMisc;rPlanner()">${ui.editMisc ? "Done" : "Edit list"}</button></div>
     <p class="hint" style="margin:-4px 0 10px">${ui.editMisc ? "Tap an item to delete it." : "Tap to highlight what you need this week (yellow = goes on the list)."}</p>
     <div class="chips">${misc}</div>
     <div class="row2" style="margin-top:12px"><input type="text" id="misc-new" placeholder="Add an extra item…" onkeydown="if(event.key==='Enter')addMisc()"><button class="btn small ghost" style="flex:0 0 auto" onclick="addMisc()">Add</button></div>
     <div class="secthead"><h2>Done planning?</h2></div>
     <div class="stack"><button class="btn ghost" onclick="newWeek()">Start a new week (clear meals &amp; ticks)</button></div>
     <div class="sticky"><button class="btn" onclick="populate()">Complete – populate shopping list</button></div>
     <div class="syncline"></div>`;
  state.meals.forEach((_, i) => { renderSugs(i); renderIngChips(i); });
}

function toggleMeal(i) { ui.openMeals.has(i) ? ui.openMeals.delete(i) : ui.openMeals.add(i); const m = $("meal-" + i); if (m) m.classList.toggle("open", ui.openMeals.has(i)); }
function editMeal(i, field, val) {
  state.meals[i][field] = val; touch();
  const t = $("mt-" + i), s = $("ms-" + i), card = $("meal-" + i), m = state.meals[i];
  if (t) t.textContent = m.name.trim() || "Meal " + (i + 1);
  if (s) s.textContent = mealSummary(m);
  if (card) card.classList.toggle("filled", !!m.name.trim());
  if (field === "ingredients") renderIngChips(i);
}
const ingLines = (s) => String(s || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
// Show each ingredient as a chip; tap one to remove it (handy for auto-filled ingredients you don't need this week).
function renderIngChips(i) {
  const box = $("ic-" + i); if (!box) return;
  const list = ingLines(state.meals[i].ingredients), n = $("icn-" + i);
  box.innerHTML = list.map((g, gi) => `<button class="chip" onclick="removeIng(${i},${gi})">${esc(g)} <span class="x">✕</span></button>`).join("");
  if (n) n.textContent = list.length ? "— tap ✕ to remove one" : "— one per line";
}
function setReuse(i, on) {
  state.meals[i].reuse = !!on; touch();
  const s = $("ms-" + i); if (s) s.textContent = mealSummary(state.meals[i]);
  const c = $("meal-" + i); if (c) c.classList.toggle("reuse", !!on);
}
function removeIng(i, gi) {
  const list = ingLines(state.meals[i].ingredients); list.splice(gi, 1);
  const v = list.join("\n"), ta = $("mi-" + i); if (ta) ta.value = v;
  editMeal(i, "ingredients", v); renderSugs(i);
}
function clearMeal(i) { if (!confirm(`Clear meal ${i + 1}?`)) return; state.meals[i] = emptyMeal(); touch(); rPlanner(); }
function toggleMisc(i) { state.misc[i].selected = !state.misc[i].selected; touch(); rPlanner(); }
function delMisc(i) { state.misc.splice(i, 1); touch(); rPlanner(); }
function addMisc() {
  const inp = $("misc-new"), v = inp.value.trim(); if (!v) return;
  state.misc.push({ id: "m" + Date.now(), name: v, selected: true }); touch(); rPlanner(); toast(`Added "${v}" (highlighted)`);
}
function newWeek() {
  if (!confirm("Clear all meals and un-highlight the extras? Your aisles and extras list stay.")) return;
  const left = state.meals.filter((m) => m.name.trim() || String(m.ingredients || "").trim());
  const keep = left.length && confirm(`You have ${left.length} meal${left.length === 1 ? "" : "s"} you haven't cooked yet.\n\nOK = keep ${left.length === 1 ? "it" : "them"} in the new week as "re-used" (ingredients left off the shopping list).\nCancel = clear ${left.length === 1 ? "it" : "them"} too.`);
  state.meals = Array.from({ length: MAX_MEALS }, emptyMeal); state.usedMeals = [];
  if (keep) left.forEach((m, k) => { state.meals[k] = { ...m, reuse: true }; });
  state.misc.forEach((x) => (x.selected = false));
  ui.openMeals = new Set([0]); ui.auto = {}; ui.autoNote = {}; touch(); rPlanner(); toast("Fresh week ready ✓");
}

function populate() {
  const items = buildShopping(state);
  if (!items.length) { alert("No ingredients found! Add some meals or tick some extras first."); return; }
  state.shopping = { generatedAt: new Date().toISOString(), items };
  const rec = recordMeals(state);       // remember these meals + ingredients for next time
  touch();
  backupSoon();
  const unk = items.filter((i) => !i.aisle).length;
  ui.shopBanner = `✅ Shopping list populated — ${items.length} unique items sorted by aisle.${unk ? ` ${unk} need an aisle.` : ""}${rec.meals ? ` ${rec.meals} meal${rec.meals === 1 ? "" : "s"} saved to Previous meals${rec.added ? ` (+${rec.added} new ingredient${rec.added === 1 ? "" : "s"})` : ""}.` : ""}`;
  go("shopping");
}

/* ---------------------------------------------------------------------
   WEEKLY RECIPES
   --------------------------------------------------------------------- */
function recipeCard(m, kind, i) {
  const ing = String(m.ingredients).split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const where = [m.book && m.book.trim(), m.page && String(m.page).trim() ? "page " + String(m.page).trim() : ""].filter(Boolean).join(" · ");
  const used = kind === "used";
  return `<div class="card rcp ${used ? "used" : ""}"><h3>${used ? "✓ " : ""}${esc(m.name.trim() || "Untitled meal")}${m.reuse && !used ? ' <span class="tag">♻ re-used</span>' : ""}</h3><div class="bk">${esc(where || "No recipe book noted")}</div>${ing.length ? `<ul>${ing.map((l) => `<li>${esc(l)}</li>`).join("")}</ul>` : '<div class="hint">No ingredients listed.</div>'}
    <div class="row2" style="margin-top:10px"><button class="btn small ${used ? "ghost" : ""}" onclick="${used ? "unuseMeal" : "toggleUsed"}(${i})">${used ? "Undo – not used" : "Mark as used"}</button><button class="btn small danger" onclick="${used ? "removeUsed" : "removeRecipe"}(${i})">Remove</button></div></div>`;
}
function rRecipes() {
  const idx = state.meals.map((m, i) => i).filter((i) => state.meals[i].name.trim() || String(state.meals[i].ingredients).trim());
  const done = state.usedMeals, total = idx.length + done.length;
  $("v-recipes").innerHTML = bar("Weekly recipes") + (total
    ? `<p class="sub">${total} meal${total === 1 ? "" : "s"} this week${done.length ? ` · ${done.length} used, ${idx.length} to go` : ""}. <b>Mark as used</b> once you've cooked one — it's struck off here and cleared from the Meal planner.</p>` +
      idx.map((i) => recipeCard(state.meals[i], "todo", i)).join("") +
      (done.length ? `<div class="secthead"><h2>Used this week</h2></div>` + done.map((m, j) => recipeCard(m, "used", j)).join("") : "")
    : `<div class="empty"><div class="big">📖</div><b>No meals planned yet</b><p>Add them in the Meal planner and they'll show up here.</p><button class="btn" onclick="go('planner')">Go to Meal planner</button></div>`) +
    `<div class="syncline"></div>`;
}
// Cooked it: strike it off here, clear its slot in the Meal planner, and back up. The shopping list is untouched.
function toggleUsed(i) {
  const m = state.meals[i];
  state.usedMeals.push({ name: m.name, book: m.book, page: m.page, ingredients: m.ingredients });
  state.meals[i] = emptyMeal(); delete ui.auto[i]; delete ui.autoNote[i];
  touch(); backupSoon(); rRecipes();
}
// Undo: put it back in the first empty planner slot.
function unuseMeal(j) {
  const i = state.meals.findIndex((m) => !m.name.trim() && !String(m.ingredients || "").trim());
  if (i < 0) return alert(`All ${MAX_MEALS} meal slots in the Meal planner are full — clear one first.`);
  state.meals[i] = state.usedMeals.splice(j, 1)[0]; touch(); backupSoon(); rRecipes();
}
function removeUsed(j) {
  if (!confirm(`Remove "${state.usedMeals[j].name.trim() || "this meal"}" from this week's list?`)) return;
  state.usedMeals.splice(j, 1); touch(); backupSoon(); rRecipes();
}
function removeRecipe(i) {
  if (!confirm(`Remove "${state.meals[i].name.trim() || "this meal"}" from this week's plan? Its saved ingredients in Previous meals are kept.`)) return;
  state.meals[i] = emptyMeal(); touch(); backupSoon(); rRecipes();
}

/* ---------------------------------------------------------------------
   AISLE PLANNER
   --------------------------------------------------------------------- */
function addKeyword(k, aisle) {
  k = cleanItem(k); if (!k) return false;
  const ex = state.keywords.find((x) => x.k === k);
  if (ex) ex.a = aisle; else state.keywords.push({ k, a: aisle });
  return true;
}
function reclassify() { if (state.shopping) state.shopping.items.forEach((it) => (it.aisle = findAisle(it.name, state))); }

function rAisles() {
  const unknown = state.shopping ? state.shopping.items.map((it, i) => ({ it, i })).filter((x) => !x.it.aisle || !state.aisles.includes(x.it.aisle)) : [];
  const counts = {}; state.keywords.forEach((k) => (counts[k.a] = (counts[k.a] || 0) + 1));
  $("v-aisles").innerHTML = bar("Aisle planner") +
    `<p class="sub">Set the order you walk the shop. The shopping list follows this order.</p>
     <div class="card"><h3>Aisle order</h3>${state.aisles.map((a, i) => `
       <div class="arow"><span class="ord">${i + 1}</span><span class="an">${esc(a)}<small>${counts[a] || 0} keyword${(counts[a] || 0) === 1 ? "" : "s"}</small></span>
       <button class="ib" aria-label="Move up" onclick="moveAisle(${i},-1)" ${i === 0 ? "disabled" : ""}>▲</button>
       <button class="ib" aria-label="Move down" onclick="moveAisle(${i},1)" ${i === state.aisles.length - 1 ? "disabled" : ""}>▼</button>
       <button class="ib" aria-label="Rename" onclick="renameAisle(${i})">✎</button>
       <button class="ib" aria-label="Delete" onclick="delAisle(${i})">✕</button></div>`).join("")}
       <div class="row2" style="margin-top:12px"><input type="text" id="aisle-new" placeholder="New aisle name…"><button class="btn small ghost" style="flex:0 0 auto" onclick="addAisle()">Add aisle</button></div></div>
     <div class="card"><h3>Unknown ingredients ${unknown.length ? `(${unknown.length})` : ""}</h3>
       ${unknown.length ? unknown.map(({ it, i }) => `<div class="unkrow"><label class="l" style="margin-top:0">Keyword to remember</label>
         <input type="text" id="uk-${i}" value="${esc(it.name)}"><div class="row2" style="margin-top:8px"><select id="ua-${i}"><option value="">Choose aisle…</option>${aisleOptions("")}</select><button class="btn small" style="flex:0 0 auto" onclick="saveUnknown(${i})">Save</button></div></div>`).join("")
         : '<p class="hint" style="margin:0">Nothing unknown on your current list. Items the app can\'t place will appear here.</p>'}</div>
     <div class="card"><h3>Add a keyword manually</h3>
       <p class="hint" style="margin:0 0 4px">If an ingredient <i>contains</i> the keyword, it goes in that aisle. Longer, more specific keywords win.</p>
       <label class="l">Keyword</label><input type="text" id="kw-new" placeholder="e.g. green beans">
       <label class="l">Aisle</label><select id="kw-aisle">${aisleOptions("")}</select>
       <button class="btn" style="margin-top:12px" onclick="addKw()">Add keyword</button></div>
     <div class="card"><h3>All keywords (${state.keywords.length})</h3>
       <input type="text" id="kw-filter" placeholder="Search keywords…" value="${esc(ui.kwFilter)}" oninput="ui.kwFilter=this.value;ui.kwOpen=true;renderKwList()">
       <div id="kw-list"></div></div><div class="syncline"></div>`;
  renderKwList();
}

function renderKwList() {
  const box = $("kw-list"); if (!box) return;
  const f = cleanItem(ui.kwFilter);
  if (!f && !ui.kwOpen) { box.innerHTML = `<button class="btn ghost" style="margin-top:10px" onclick="ui.kwOpen=true;renderKwList()">Show all ${state.keywords.length}</button>`; return; }
  const rows = state.keywords.map((k, i) => ({ k, i })).filter((x) => !f || x.k.k.includes(f) || x.k.a.toLowerCase().includes(f))
    .sort((a, b) => aisleRank(a.k.a, state) - aisleRank(b.k.a, state) || (a.k.k < b.k.k ? -1 : 1));
  box.innerHTML = rows.length ? rows.map(({ k, i }) => `<div class="krow"><span class="kk">${esc(k.k)}</span>
    <select onchange="setKwAisle(${i},this.value)">${aisleOptions(k.a)}${state.aisles.includes(k.a) ? "" : `<option selected>${esc(k.a)}</option>`}</select>
    <button class="ib" aria-label="Delete keyword" onclick="delKw(${i})">✕</button></div>`).join("") : '<p class="hint">No keywords match.</p>';
}

function moveAisle(i, d) { const j = i + d; if (j < 0 || j >= state.aisles.length) return; [state.aisles[i], state.aisles[j]] = [state.aisles[j], state.aisles[i]]; touch(); backupSoon(); rAisles(); }
function renameAisle(i) {
  const old = state.aisles[i], nv = (prompt("Rename aisle:", old) || "").trim(); if (!nv || nv === old) return;
  if (state.aisles.some((a) => a.toLowerCase() === nv.toLowerCase())) return alert("An aisle with that name already exists.");
  state.aisles[i] = nv; state.keywords.forEach((k) => { if (k.a === old) k.a = nv; });
  if (state.shopping) state.shopping.items.forEach((it) => { if (it.aisle === old) it.aisle = nv; });
  touch(); rAisles();
}
function delAisle(i) {
  const a = state.aisles[i], n = state.keywords.filter((k) => k.a === a).length;
  if (n) return alert(`"${a}" still has ${n} keyword${n === 1 ? "" : "s"}. Move or delete them first (search "${a}" in All keywords).`);
  if (!confirm(`Delete the aisle "${a}"?`)) return;
  state.aisles.splice(i, 1); reclassify(); touch(); backupSoon(); rAisles();
}
function addAisle() {
  const v = $("aisle-new").value.trim(); if (!v) return;
  if (state.aisles.some((a) => a.toLowerCase() === v.toLowerCase())) return alert("That aisle already exists.");
  state.aisles.push(v); touch(); rAisles(); toast(`Added aisle "${v}"`);
}
function addKw() {
  const k = $("kw-new").value, a = $("kw-aisle").value;
  if (!cleanItem(k)) return alert("Type a keyword first."); if (!a) return alert("Choose an aisle.");
  addKeyword(k, a); reclassify(); touch(); rAisles(); toast("Keyword saved");
}
function saveUnknown(i) {
  const k = $("uk-" + i).value, a = $("ua-" + i).value;
  if (!cleanItem(k)) return alert("Keyword can't be empty."); if (!a) return alert("Choose an aisle first.");
  addKeyword(k, a); reclassify(); touch(); rAisles(); toast("Remembered ✓");
}
function setKwAisle(i, a) { state.keywords[i].a = a; reclassify(); touch(); }
function delKw(i) { state.keywords.splice(i, 1); reclassify(); touch(); renderKwList(); }

/* ---------------------------------------------------------------------
   SETTINGS
   --------------------------------------------------------------------- */
function rSettings() {
  const g = getSettings().github;
  const msg = ui.settingsMsg; ui.settingsMsg = null;
  $("v-settings").innerHTML = bar("Settings") + (msg ? `<div class="ok">${esc(msg)}</div>` : "") +
    `<div class="card"><h3>Automatic backup (GitHub)</h3>
      <label class="l">Personal access token</label><input type="password" id="gh-token" value="${esc(g.token)}" placeholder="github_pat_…" autocomplete="off">
      <div class="row2"><div><label class="l">GitHub username</label><input type="text" id="gh-owner" value="${esc(g.owner)}" autocapitalize="off" autocorrect="off"></div>
      <div><label class="l">Repo name</label><input type="text" id="gh-repo" value="${esc(g.repo)}" autocapitalize="off" autocorrect="off"></div></div>
      <label class="l">File path in repo</label><input type="text" id="gh-path" value="${esc(g.path)}" autocapitalize="off">
      <div class="switch"><span>Back up automatically after changes</span><input type="checkbox" id="gh-auto" ${g.auto ? "checked" : ""}></div>
      <p class="hint">Use a <b>separate private repo</b> for your data (not the public one hosting this app) and a fine-grained token limited to that repo with <b>Contents: Read and write</b>.</p>
      <div class="stack" style="margin-top:12px"><button class="btn" onclick="saveSettingsForm()">Save</button>
      <button class="btn ghost" onclick="manualBackup()">Back up now</button>
      <button class="btn ghost" onclick="loadSaved()">Load saved data</button></div>
      <div id="load-msg" class="hint" style="margin-top:10px"></div><div class="syncline"></div></div>
     <div class="card"><h3>Settings backup</h3>
      <p class="hint" style="margin:0 0 8px">Your token lives only in this app's storage on this phone. Keep a copy in Notes or a password manager so a wipe is a one-paste fix.</p>
      <button class="btn ghost" onclick="showSettingsText()">Show settings text</button>
      <textarea id="set-out" readonly style="display:none;margin-top:10px;min-height:90px;font-size:12px" onclick="this.select()"></textarea>
      <p class="hint" id="set-hint" style="display:none">Tap the text, choose Select All → Copy, then paste into Notes.</p>
      <label class="l">Restore from saved text</label><textarea id="set-in" style="min-height:80px;font-size:12px" placeholder="Paste your saved settings text here"></textarea>
      <button class="btn ghost" style="margin-top:10px" onclick="restoreSettingsText()">Restore settings</button><div id="set-msg" class="hint" style="margin-top:8px"></div></div>
     <div class="card"><h3>Start again</h3><p class="hint" style="margin:0 0 10px">Replaces the aisles, keywords, extras and meals with the original spreadsheet data. Your Previous meals history is kept.</p>
      <button class="btn danger" onclick="resetAll()">Reset to spreadsheet defaults</button></div>`;
}
function readGh() { return { token: $("gh-token").value.trim(), owner: $("gh-owner").value.trim(), repo: $("gh-repo").value.trim(), path: $("gh-path").value.trim() || "meal-data.json", auto: $("gh-auto").checked }; }
function saveSettingsForm() { putSettings({ github: readGh() }); $("load-msg").className = "ok"; $("load-msg").textContent = "✅ Saved"; renderSyncLines(); }
async function manualBackup() {
  putSettings({ github: readGh() });
  const m = $("load-msg"); m.className = "hint"; m.textContent = "Backing up…";
  const r = await backupNow();
  m.className = r.ok ? "ok" : "err"; m.textContent = r.ok ? `✅ Backed up successfully — ${new Date().toLocaleTimeString()}` : "Backup failed: " + (r.error || "kept the GitHub copy");
}
function showSettingsText() {
  const t = JSON.stringify(getSettings().github, null, 1), out = $("set-out");
  out.style.display = "block"; $("set-hint").style.display = "block"; out.value = t; out.focus(); out.setSelectionRange(0, t.length);
  if (navigator.clipboard) navigator.clipboard.writeText(t).catch(() => {});
}
function restoreSettingsText() {
  const m = $("set-msg");
  try { const p = JSON.parse($("set-in").value.trim()); putSettings({ github: { ...defaultSettings().github, ...p } }); m.className = "ok"; m.textContent = "✅ Restored"; setTimeout(rSettings, 600); }
  catch (e) { m.className = "err"; m.textContent = "That doesn't look right — paste the whole text."; }
}
function resetAll() {
  if (!confirm("Reset EVERYTHING on this phone to the original spreadsheet data? Meals, extras, aisles and keywords will be replaced. Your Previous meals history is kept.")) return;
  const keep = state.history; state = freshState(); normalise(); state.history = keep; touch(); ui.settingsMsg = "Reset to spreadsheet defaults (your Previous meals were kept)."; go("settings");
}

/* ---------------------------------------------------------------------
   PLANNER: recognise previous meals and fill in their ingredients
   --------------------------------------------------------------------- */
const AUTO_NOTE = "Filled in from Previous meals — delete anything you don't need this week.";

// Safe to overwrite the ingredients box if it's empty, or still exactly what we filled in earlier.
function canFillIngredients(i) {
  const m = state.meals[i], cur = String(m.ingredients || "").trim();
  if (!cur) return true;
  const a = ui.auto[i];
  return !!(a && m.ingredients === a.text);
}

// Update the meal's on-screen fields in place (no full redraw, so the keyboard/focus isn't disturbed).
function refreshMealFields(i) {
  const m = state.meals[i], set = (id, v) => { const el = $(id); if (el) el.value = v; };
  const rc = $("mr-" + i); if (rc) rc.checked = !!m.reuse;
  set("mn-" + i, m.name); set("mb-" + i, m.book); set("mp-" + i, m.page); set("mi-" + i, m.ingredients);
  const t = $("mt-" + i), s = $("ms-" + i), c = $("meal-" + i), a = $("auto-" + i);
  if (t) t.textContent = m.name.trim() || "Meal " + (i + 1);
  if (s) s.textContent = mealSummary(m);
  if (c) c.classList.toggle("filled", !!m.name.trim());
  if (a) a.textContent = ui.autoNote[i] ? AUTO_NOTE : "";
  renderSugs(i); renderIngChips(i);
}

function fillFromHistory(i, h, { override = false } = {}) {
  const m = state.meals[i];
  if (override) { m.name = h.name; if (h.book) m.book = h.book; if (h.page) m.page = h.page; }
  else { if (!m.book.trim() && h.book) m.book = h.book; if (!String(m.page).trim() && h.page) m.page = h.page; }
  m.ingredients = h.ingredients.join("\n");
  ui.auto[i] = { key: mealKey(h.name), text: m.ingredients };
  ui.autoNote[i] = h.ingredients.length > 0;
  touch(); refreshMealFields(i);
}

function onNameInput(i, val) { editMeal(i, "name", val); renderSugs(i); }

// When you finish typing a name that matches a previous meal, fill in its saved ingredients.
function onNameCommit(i) {
  const m = state.meals[i], h = findHistory(state, m.name);
  if (!h || !h.ingredients.length) return;
  if (canFillIngredients(i) && m.ingredients !== h.ingredients.join("\n")) fillFromHistory(i, h);
}

function pickHistory(i, id) {
  const h = state.history.find((x) => x.id === id); if (!h) return;
  if (!canFillIngredients(i) && !confirm("Replace the ingredients already in this meal with the saved ones?")) return;
  fillFromHistory(i, h, { override: true });
}

function renderSugs(i) {
  const box = $("sug-" + i); if (!box) return;
  if (!state.history.length) { box.innerHTML = ""; return; }
  const m = state.meals[i], q = mealKey(m.name), exact = findHistory(state, m.name);
  let list;
  if (!q) list = top5(state);
  else list = state.history.filter((h) => mealKey(h.name).includes(q) && h !== exact).sort((a, b) => b.count - a.count).slice(0, 5);
  let html = "";
  if (exact && exact.ingredients.length && m.ingredients !== exact.ingredients.join("\n"))
    html += `<button class="sug" onclick="pickHistory(${i},'${exact.id}')">↻ Load saved ingredients</button>`;
  if (!q && list.length) html += `<span class="hint" style="width:100%;margin:0">From previous meals:</span>`;
  html += list.map((h) => `<button class="sug" onclick="pickHistory(${i},'${h.id}')">${esc(h.name)} <small>×${h.count}</small></button>`).join("");
  box.innerHTML = html;
}

/* ---------------------------------------------------------------------
   PREVIOUS MEALS
   --------------------------------------------------------------------- */
const whereText = (h) => [h.book, h.page ? "p. " + h.page : ""].filter(Boolean).join(" · ") || "No book noted";

function rHistory() {
  const h = state.history;
  $("v-history").innerHTML = bar("Previous meals") + (h.length
    ? `<p class="sub">Every meal you've planned, with all the ingredients you've used for it. It grows each time you populate the shopping list.</p>
       <div class="card"><h3>🏆 Top 5 recipes</h3>${top5(state).map((x, i) => `<div class="toprow"><span class="rk">${i + 1}</span><span class="tn"><b>${esc(x.name)}</b><small>${esc(whereText(x))}</small></span><span class="cnt">×${x.count}</span></div>`).join("")}</div>
       <div class="secthead"><h2>All meals (${h.length})</h2></div>
       <input type="text" id="hist-filter" placeholder="Search meals or ingredients…" value="${esc(ui.histFilter)}" oninput="ui.histFilter=this.value;renderHistList()" autocomplete="off">
       <div id="hist-list" style="margin-top:12px"></div>`
    : `<div class="empty"><div class="big">🕘</div><b>Nothing here yet</b><p>Meals are saved automatically each time you tap <b>Complete – populate shopping list</b>, and their ingredients build up over time.</p><button class="btn" onclick="go('planner')">Go to Meal planner</button></div>`) +
    `<div class="syncline"></div>`;
  renderHistList();
}

function renderHistList() {
  const box = $("hist-list"); if (!box) return;
  const f = mealKey(ui.histFilter);
  const rows = state.history.filter((h) => !f || mealKey(h.name).includes(f) || h.ingredients.some((g) => g.toLowerCase().includes(f)) || (h.book || "").toLowerCase().includes(f))
    .sort((a, b) => String(b.lastUsed).localeCompare(String(a.lastUsed)) || a.name.localeCompare(b.name));
  box.innerHTML = rows.length ? rows.map((h) => `
    <div class="meal filled ${ui.openHist.has(h.id) ? "open" : ""}" id="hm-${h.id}">
      <button class="mhead" onclick="toggleHist('${h.id}')"><span class="mnum">${h.count}</span>
        <span class="mtxt"><span class="mt">${esc(h.name)}</span><span class="ms">${esc(whereText(h))} · ${h.ingredients.length} ingredient${h.ingredients.length === 1 ? "" : "s"}</span></span><span class="chev">›</span></button>
      <div class="mbody">
        <p class="hint" style="margin:0 0 4px">Chosen ${h.count} time${h.count === 1 ? "" : "s"}${h.lastUsed ? " · last " + new Date(h.lastUsed).toLocaleDateString([], { day: "numeric", month: "short", year: "numeric" }) : ""}</p>
        <div class="row2"><div><label class="l">Recipe book</label><input type="text" value="${esc(h.book)}" onchange="histField('${h.id}','book',this.value)"></div>
        <div class="narrow"><label class="l">Page</label><input type="text" inputmode="numeric" value="${esc(h.page)}" onchange="histField('${h.id}','page',this.value)"></div></div>
        <label class="l">Saved ingredients — tap one to remove it</label>
        <div class="chips">${h.ingredients.length ? h.ingredients.map((g, gi) => `<button class="chip" onclick="histDelIng('${h.id}',${gi})">${esc(g)} <span class="x">✕</span></button>`).join("") : '<span class="hint">None yet.</span>'}</div>
        <div class="row2" style="margin-top:10px"><input type="text" id="hi-${h.id}" placeholder="Add an ingredient…" autocomplete="off" onkeydown="if(event.key==='Enter')histAddIng('${h.id}')"><button class="btn small ghost" style="flex:0 0 auto" onclick="histAddIng('${h.id}')">Add</button></div>
        <div class="stack" style="margin-top:12px"><button class="btn" onclick="useMeal('${h.id}')">Add to this week's plan</button><button class="btn danger" onclick="histDelete('${h.id}')">Delete this meal</button></div>
      </div></div>`).join("") : '<p class="hint">No meals match.</p>';
}

function toggleHist(id) { ui.openHist.has(id) ? ui.openHist.delete(id) : ui.openHist.add(id); const el = $("hm-" + id); if (el) el.classList.toggle("open", ui.openHist.has(id)); }
function histField(id, field, val) { const h = state.history.find((x) => x.id === id); if (!h) return; h[field] = val.trim(); touch(); renderHistList(); }
function histDelIng(id, gi) { const h = state.history.find((x) => x.id === id); if (!h) return; h.ingredients.splice(gi, 1); touch(); renderHistList(); }
function histAddIng(id) {
  const h = state.history.find((x) => x.id === id), inp = $("hi-" + id); if (!h || !inp) return;
  const k = cleanItem(inp.value); if (!k) return;
  if (h.ingredients.some((g) => cleanItem(g) === k)) { toast("Already saved for this meal"); return; }
  h.ingredients.push(ingDisplay(inp.value)); touch(); renderHistList(); toast("Ingredient added");
}
function histDelete(id) {
  const h = state.history.find((x) => x.id === id); if (!h || !confirm(`Delete "${h.name}" from Previous meals? Its saved ingredients and count will be lost.`)) return;
  state.history = state.history.filter((x) => x.id !== id); touch(); rHistory();
}
// Drop a previous meal into the first empty slot of this week's plan, with its saved ingredients.
function useMeal(id) {
  const h = state.history.find((x) => x.id === id); if (!h) return;
  if (state.meals.some((m) => mealKey(m.name) === mealKey(h.name))) return alert(`"${h.name}" is already in this week's plan.`);
  const i = state.meals.findIndex((m) => !m.name.trim() && !String(m.ingredients || "").trim());
  if (i < 0) return alert(`All ${MAX_MEALS} meal slots are full — clear one in the Meal planner first.`);
  ui.openMeals = new Set([i]); fillFromHistory(i, h, { override: true }); go("planner"); toast(`Added to meal ${i + 1}`);
}

/* ---------------------------------------------------------------------
   Start-up
   --------------------------------------------------------------------- */
loadState();
render();
if ("serviceWorker" in navigator) window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
syncOnOpen();
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") syncOnOpen();
  else if (ghCfg() && getSettings().github.auto && state.updatedAt && state.updatedAt !== getSynced()) { clearTimeout(backupTimer); backupNow({ quiet: true }); }
});
