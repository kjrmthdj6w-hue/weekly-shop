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
  st.meals.forEach((m) => String(m.ingredients || "").split(/\r?\n/).forEach((l) => out.push(l)));
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
let state;
const ui = { screen: "home", openMeals: new Set([0]), editMisc: false, shopBanner: "", kwFilter: "", kwOpen: false, settingsMsg: null };

function emptyMeal() { return { name: "", book: "", page: "", ingredients: "" }; }

function freshState() {
  const s = JSON.parse(JSON.stringify(SEED));
  s.version = 1; s.updatedAt = null; s.shopping = null;
  return s;
}

function normalise() {
  state.aisles = state.aisles || [];
  state.keywords = state.keywords || [];
  state.misc = state.misc || [];
  state.meals = (state.meals || []).map((m) => ({ ...emptyMeal(), ...m }));
  while (state.meals.length < MAX_MEALS) state.meals.push(emptyMeal());
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
function defaultSettings() { return { github: { token: "", owner: "", repo: "", path: "meal-data.json", auto: true } }; }
function getSettings() {
  const d = defaultSettings();
  try { const p = JSON.parse(localStorage.getItem(LS_SETTINGS)); if (p && p.github) d.github = { ...d.github, ...p.github }; } catch (e) {}
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
const SCREENS = ["home", "shopping", "planner", "aisles", "recipes", "settings"];
function go(name) {
  ui.screen = name;
  SCREENS.forEach((s) => $("s-" + s).classList.toggle("active", s === name));
  render();
  window.scrollTo(0, 0);
}
function render() {
  ({ home: rHome, shopping: rShopping, planner: rPlanner, aisles: rAisles, recipes: rRecipes, settings: rSettings })[ui.screen]();
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
    <div class="tiles">
      <button class="tile main" onclick="go('shopping')"><span class="ic">🛒</span><span><div class="tt">Shopping list</div><div class="ts">${sh ? `${n} items · ${ticked} ticked` : "Not populated yet"}</div></span></button>
      <button class="tile" onclick="go('planner')"><span class="ic">🍽️</span><span><div class="tt">Meal planner</div><div class="ts">${planned} of ${MAX_MEALS} meals planned</div></span></button>
      <button class="tile" onclick="go('aisles')"><span class="ic">🧭</span><span><div class="tt">Aisle planner</div><div class="ts">${state.aisles.length} aisles · ${state.keywords.length} keywords</div></span></button>
      <button class="tile" onclick="go('recipes')"><span class="ic">📖</span><span><div class="tt">Weekly recipes</div><div class="ts">${planned} recipe${planned === 1 ? "" : "s"} this week</div></span></button>
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
  return [bits, n ? `${n} ingredient${n === 1 ? "" : "s"}` : ""].filter(Boolean).join(" · ") || "Tap to add";
}

function rPlanner() {
  const books = [...new Set(state.meals.map((m) => (m.book || "").trim()).filter(Boolean))];
  const meals = state.meals.slice(0, Math.max(MAX_MEALS, state.meals.length)).map((m, i) => `
    <div class="meal ${m.name.trim() ? "filled" : ""} ${ui.openMeals.has(i) ? "open" : ""}" id="meal-${i}">
      <button class="mhead" onclick="toggleMeal(${i})"><span class="mnum">${i + 1}</span>
        <span class="mtxt"><span class="mt" id="mt-${i}">${esc(m.name.trim() || "Meal " + (i + 1))}</span><span class="ms" id="ms-${i}">${esc(mealSummary(m))}</span></span><span class="chev">›</span></button>
      <div class="mbody">
        <label class="l">Meal name</label><input type="text" value="${esc(m.name)}" oninput="editMeal(${i},'name',this.value)" placeholder="e.g. Greek salad">
        <div class="row2"><div><label class="l">Recipe book</label><input type="text" list="books" value="${esc(m.book)}" oninput="editMeal(${i},'book',this.value)" placeholder="e.g. Fast food"></div>
        <div class="narrow"><label class="l">Page</label><input type="text" inputmode="numeric" value="${esc(m.page)}" oninput="editMeal(${i},'page',this.value)" placeholder="18"></div></div>
        <label class="l">Ingredients — one per line</label>
        <textarea oninput="editMeal(${i},'ingredients',this.value)" placeholder="cherry tomatoes&#10;cucumber&#10;feta">${esc(m.ingredients)}</textarea>
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
}

function toggleMeal(i) { ui.openMeals.has(i) ? ui.openMeals.delete(i) : ui.openMeals.add(i); const m = $("meal-" + i); if (m) m.classList.toggle("open", ui.openMeals.has(i)); }
function editMeal(i, field, val) {
  state.meals[i][field] = val; touch();
  const t = $("mt-" + i), s = $("ms-" + i), card = $("meal-" + i), m = state.meals[i];
  if (t) t.textContent = m.name.trim() || "Meal " + (i + 1);
  if (s) s.textContent = mealSummary(m);
  if (card) card.classList.toggle("filled", !!m.name.trim());
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
  state.meals = Array.from({ length: MAX_MEALS }, emptyMeal);
  state.misc.forEach((x) => (x.selected = false));
  ui.openMeals = new Set([0]); touch(); rPlanner(); toast("Fresh week ready ✓");
}

function populate() {
  const items = buildShopping(state);
  if (!items.length) { alert("No ingredients found! Add some meals or tick some extras first."); return; }
  state.shopping = { generatedAt: new Date().toISOString(), items };
  touch();
  if (ghCfg() && getSettings().github.auto) { clearTimeout(backupTimer); backupNow({ quiet: true }); }
  const unk = items.filter((i) => !i.aisle).length;
  ui.shopBanner = `✅ Shopping list populated — ${items.length} unique items sorted by aisle.${unk ? ` ${unk} need an aisle.` : ""}`;
  go("shopping");
}

/* ---------------------------------------------------------------------
   WEEKLY RECIPES
   --------------------------------------------------------------------- */
function rRecipes() {
  const meals = state.meals.filter((m) => m.name.trim() || String(m.ingredients).trim());
  $("v-recipes").innerHTML = bar("Weekly recipes") + (meals.length
    ? `<p class="sub">This week's ${meals.length} meal${meals.length === 1 ? "" : "s"} and what each needs. Edit them in the Meal planner.</p>` +
      meals.map((m) => {
        const ing = String(m.ingredients).split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
        const where = [m.book && m.book.trim(), m.page && String(m.page).trim() ? "page " + String(m.page).trim() : ""].filter(Boolean).join(" · ");
        return `<div class="card rcp"><h3>${esc(m.name.trim() || "Untitled meal")}</h3><div class="bk">${esc(where || "No recipe book noted")}</div>${ing.length ? `<ul>${ing.map((l) => `<li>${esc(l)}</li>`).join("")}</ul>` : '<div class="hint">No ingredients listed.</div>'}</div>`;
      }).join("")
    : `<div class="empty"><div class="big">📖</div><b>No meals planned yet</b><p>Add them in the Meal planner and they'll show up here.</p><button class="btn" onclick="go('planner')">Go to Meal planner</button></div>`) +
    `<div class="syncline"></div>`;
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

function moveAisle(i, d) { const j = i + d; if (j < 0 || j >= state.aisles.length) return; [state.aisles[i], state.aisles[j]] = [state.aisles[j], state.aisles[i]]; touch(); rAisles(); }
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
  state.aisles.splice(i, 1); reclassify(); touch(); rAisles();
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
     <div class="card"><h3>Start again</h3><p class="hint" style="margin:0 0 10px">Replaces everything on this phone with the original spreadsheet data (aisles, keywords, extras and meals).</p>
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
  if (!confirm("Reset EVERYTHING on this phone to the original spreadsheet data? Meals, extras, aisles and keywords will be replaced.")) return;
  state = freshState(); normalise(); touch(); ui.settingsMsg = "Reset to spreadsheet defaults."; go("settings");
}

/* ---------------------------------------------------------------------
   Start-up
   --------------------------------------------------------------------- */
loadState();
render();
if ("serviceWorker" in navigator) window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
syncOnOpen();
