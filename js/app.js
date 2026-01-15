/* ============================================================================
 * 3PI BOOKED DASHBOARD – FRONTEND (Azure Static Web Apps)
 *
 * APIs REQUIRED:
 *   ✅ /api/getSheetNames
 *   ✅ /api/getSheetData?sheet=NAME
 *   ✅ /api/getBookedCounts
 *
 * DEFAULT VIEW:
 *   ✅ Current Month (auto)
 *   ✅ Today → End of Month
 *
 * SCORECARDS:
 *   ✅ Jan → Dec order
 *   ✅ Click month = full month view
 * ========================================================================== */

/* =========================
   API helper
   ========================= */
async function apiGet(path, label = path) {
  const clean = String(path || "").trim();
  const url = clean.startsWith("http") ? clean : `${window.location.origin}${clean}`;

  const res = await fetch(url, {
    method: "GET",
    headers: { Accept: "application/json" },
  });

  const text = await res.text();

  if (!res.ok) {
    console.warn(`[API:${label}] HTTP ${res.status}:`, text.slice(0, 500));
    throw new Error(`API error ${res.status} for ${label}`);
  }

  try {
    return JSON.parse(text);
  } catch {
    console.warn(`[API:${label}] Invalid JSON:`, text.slice(0, 500));
    throw new Error(`Invalid JSON returned for: ${label}`);
  }
}

/* =========================
   Loader overlay
   ========================= */
let pendingLoads = 0;
function showLoader() {
  pendingLoads++;
  document.getElementById("loaderOverlay")?.classList.remove("hidden");
}
function hideLoader() {
  pendingLoads = Math.max(0, pendingLoads - 1);
  if (!pendingLoads) document.getElementById("loaderOverlay")?.classList.add("hidden");
}

/* =========================
   Gate / Debounce
   ========================= */
let netBusy = false;
let debounceTimer = null;
function queueTask(fn) {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(async () => {
    if (netBusy) return;
    netBusy = true;
    showLoader();
    try {
      await fn();
    } finally {
      netBusy = false;
      hideLoader();
    }
  }, 120);
}

/* =========================
   Collapsibles
   ========================= */
function setOpen(btn, panel, open) {
  btn.classList.toggle("open", open);
  btn.setAttribute("aria-expanded", String(open));
  panel.classList.toggle("open", open);
  panel.setAttribute("aria-hidden", String(!open));
}

function initCollapsibles() {
  document.querySelectorAll(".sf4-collapse-btn").forEach((btn) => {
    const id = btn.getAttribute("aria-controls");
    const panel = document.getElementById(id);
    if (!panel) return;

    const initialOpen =
      btn.classList.contains("open") ||
      btn.getAttribute("aria-expanded") === "true";

    setOpen(btn, panel, initialOpen);

    btn.addEventListener("click", () => {
      const nowOpen = !btn.classList.contains("open");
      setOpen(btn, panel, nowOpen);

      // If MONTHLY collapses -> return to default view
      if (id === "panel-monthly" && !nowOpen) {
        forceDefaultMode();
        queueTask(() => reloadDefault());
      }
    });
  });
}

/* =========================
   DOM
   ========================= */
const els = {
  pageSize: document.getElementById("pageSize"),
  reloadBtn: document.getElementById("reloadBtn"),
  status: document.getElementById("status"),
  thead: document.getElementById("tableHead"),
  tbody: document.getElementById("tableBody"),
  emptyState: document.getElementById("emptyState"),
  pagination: document.getElementById("pagination"),
  cards: document.getElementById("cards"),
};

/* =========================
   Excel Column mapping
   ========================= */
const COL_DATE = 0;
const COL_EMAIL = 3;
const COL_NAME = 4;
const COL_START = 5;
const COL_END = 6;
const COL_BOOKED = 7;
const COL_SITE = 8;
const COL_ACCOUNT = 9;
const COL_TICKET = 10;
const COL_MID = 11;

/* Required output order */
const OUT_COLS = [
  { key: "Email", src: COL_EMAIL },
  { key: "Name", src: COL_NAME },
  { key: "Date", src: COL_DATE },
  { key: "Start Time", src: COL_START },
  { key: "End Time", src: COL_END },
  { key: "BOOKED", src: COL_BOOKED },
  { key: "Site", src: COL_SITE },
  { key: "Account", src: COL_ACCOUNT },
  { key: "Ticket", src: COL_TICKET },
  { key: "Shift4 MID", src: COL_MID },
];

/* =========================
   Timezone chip
   ========================= */
let tzInfo = { abbr: "", iana: "" };

/* =========================
   App state
   ========================= */
let monthTabs = [];
let activeMonthName = null;

let defaultMode = true;
let displayRowsCurrent = [];

let currentPage = 1;
let pageSize = (els.pageSize && parseInt(els.pageSize.value, 10)) || 15;
let sortCol = 2; // Date column
let sortDir = "asc";

let filterStartMs = null;
let filterEndMs = null;

/* Month ordering */
const MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const MONTH_RX =
  /^(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)(?:\s+\d{4})?$/i;

function isMonthSheet(name) {
  return MONTH_RX.test(String(name || "").trim());
}

function monthIndexOfSheet(name) {
  const s = String(name || "").trim().slice(0, 3).toLowerCase();
  const map = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
  return map[s] ?? null;
}

function sortMonthNamesJanToDec(list) {
  return list.slice().sort((a, b) => {
    const ia = monthIndexOfSheet(a);
    const ib = monthIndexOfSheet(b);
    if (ia == null && ib == null) return String(a).localeCompare(String(b));
    if (ia == null) return 1;
    if (ib == null) return -1;
    return ia - ib;
  });
}

/* =========================
   Boot
   ========================= */
document.addEventListener("DOMContentLoaded", () => {
  initCollapsibles();

  tzInfo = detectTimezone();
  ensureTzChip();

  els.reloadBtn?.addEventListener("click", () => {
    queueTask(async () => {
      await refreshCards();
      await reloadDefault();
    });
  });

  els.pageSize?.addEventListener("change", () => {
    pageSize = parseInt(els.pageSize.value, 10) || 15;
    currentPage = 1;
    renderTable();
  });

  // scorecard click -> full month view
  els.cards?.addEventListener("click", (e) => {
    const card = e.target.closest(".card[data-name]");
    if (!card) return;

    const name = card.getAttribute("data-name");
    if (!monthTabs.includes(name)) return;

    activeMonthName = name;
    defaultMode = false;
    setActiveCard(activeMonthName);

    queueTask(() => reloadSingleMonth(activeMonthName));
  });

  queueTask(() => loadTabsAndFirstPaint());
});

/* =========================
   Load tabs
   ========================= */
async function loadTabsAndFirstPaint() {
  // Pull sheet names
  const names = await apiGet("/api/getSheetNames", "getSheetNames").catch(() => []);
  monthTabs = Array.isArray(names) ? names.filter(isMonthSheet) : [];

  // Sort tabs Jan->Dec always
  monthTabs = sortMonthNamesJanToDec(monthTabs);

  if (!monthTabs.length) {
    setStatusLabel("No month sheets found", 0);
    return;
  }

  // pick current month
  activeMonthName = pickCurrentMonthTab(monthTabs);

  // default mode is current month today->end-of-month
  forceDefaultMode();

  // scorecards
  await refreshCards();

  // load table
  await reloadDefault();
}

function pickCurrentMonthTab(names) {
  const now = new Date();
  const currentShort = MONTHS_SHORT[now.getMonth()];
  const exact = names.find((n) => String(n).trim().toLowerCase() === currentShort.toLowerCase());
  return exact || names[0];
}

function endOfMonth(d) {
  return new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59, 999);
}

function forceDefaultMode() {
  defaultMode = true;

  if (!activeMonthName) activeMonthName = pickCurrentMonthTab(monthTabs);

  const now = new Date();
  const start = new Date(now);
  start.setHours(0, 0, 0, 0);

  filterStartMs = start.getTime();
  filterEndMs = endOfMonth(now).getTime();

  setActiveCard(activeMonthName);
}

/* =========================
   Scorecards
   ========================= */
let lastCardsJSON = "";

async function refreshCards() {
  els.cards.innerHTML = "";

  let items;
  try {
    items = await apiGet("/api/getBookedCounts", "getBookedCounts");
  } catch (err) {
    console.error("Scorecards failed:", err);
    els.cards.innerHTML = `
      <div style="padding:10px;color:#b91c1c;font-weight:900;">
        Failed to load monthly scorecards
      </div>
    `;
    return;
  }

  if (!Array.isArray(items)) {
    els.cards.innerHTML = `
      <div style="padding:10px;color:#b91c1c;font-weight:900;">
        Scorecards API returned invalid data
      </div>
    `;
    return;
  }

  // Enforce Jan->Dec order on frontend too
  items = items
    .filter((x) => x && isMonthSheet(x.name))
    .sort((a, b) => (monthIndexOfSheet(a.name) ?? 99) - (monthIndexOfSheet(b.name) ?? 99));

  const j = JSON.stringify(items);
  if (j === lastCardsJSON) {
    setActiveCard(activeMonthName);
    return;
  }
  lastCardsJSON = j;

  els.cards.innerHTML = items
    .map(({ name, count }) => `
      <div class="card" data-name="${escapeHtml(name)}" title="${escapeHtml(name)}">
        <span class="chip">BOOKED</span>
        <div class="title">${escapeHtml(name)}</div>
        <div class="value">${Number(count || 0)}</div>
      </div>
    `)
    .join("");

  setActiveCard(activeMonthName);
}

function setActiveCard(name) {
  els.cards?.querySelectorAll(".card").forEach((c) => {
    c.classList.toggle("active", c.getAttribute("data-name") === name);
  });
}

/* =========================
   Reload handlers
   ========================= */
async function reloadDefault() {
  if (!activeMonthName) return;

  resetTableState();

  const payload = await apiGet(
    `/api/getSheetData?sheet=${encodeURIComponent(activeMonthName)}`,
    `getSheetData:${activeMonthName}`
  ).catch(() => ({ rows: [] }));

  const rawRows = Array.isArray(payload.rows) ? payload.rows : [];

  // Map + filter today->end of month
  const mapped = mapRowsForDisplay(rawRows);

  const filtered = mapped.filter((r) => {
    const ms = r.__dateMs;
    if (!Number.isFinite(ms)) return false;
    return ms >= filterStartMs && ms <= filterEndMs;
  });

  displayRowsCurrent = filtered.map((x) => x.display);

  drawHeader(OUT_COLS.map((c) => c.key));
  setStatusLabel(`${activeMonthName} (Today → End of Month)`, displayRowsCurrent.length);
  renderTable();
}

async function reloadSingleMonth(tab) {
  if (!tab) return;

  resetTableState();

  const payload = await apiGet(
    `/api/getSheetData?sheet=${encodeURIComponent(tab)}`,
    `getSheetData:${tab}`
  ).catch(() => ({ rows: [] }));

  const rawRows = Array.isArray(payload.rows) ? payload.rows : [];
  const mapped = mapRowsForDisplay(rawRows);

  displayRowsCurrent = mapped.map((x) => x.display);

  drawHeader(OUT_COLS.map((c) => c.key));
  setStatusLabel(`${tab} (Full Month)`, displayRowsCurrent.length);
  renderTable();
}

/* =========================
   Row mapping + formatting
   ========================= */
function mapRowsForDisplay(rawRows) {
  return rawRows.map((row) => {
    const dateMs = excelCellToDateMs(row?.[COL_DATE]);

    const display = OUT_COLS.map((c) => {
      const v = row?.[c.src];

      if (c.key === "Date") return dateMs ? fmtLocalDate(dateMs) : "";

      if (c.key === "Start Time") {
        const ms = excelCellToTimeMs(v);
        return ms ? fmtLocalTime(ms) : "";
      }

      if (c.key === "End Time") {
        const ms = excelCellToTimeMs(v);
        return ms ? fmtLocalTime(ms) : "";
      }

      return v == null ? "" : String(v);
    });

    return { display, __dateMs: dateMs };
  });
}

function excelCellToDateMs(v) {
  if (v == null) return null;

  if (typeof v === "number" && Number.isFinite(v)) {
    return Math.round((v - 25569) * 86400 * 1000);
  }

  const s = String(v).trim();
  if (!s) return null;

  const parsed = Date.parse(s);
  if (!Number.isNaN(parsed)) return parsed;

  return null;
}

function excelCellToTimeMs(v) {
  if (v == null) return null;

  // Excel datetime number -> fractional = time
  if (typeof v === "number" && Number.isFinite(v)) {
    const frac = v - Math.floor(v);
    const totalMinutes = Math.round(frac * 24 * 60);

    const hh = Math.floor(totalMinutes / 60);
    const mm = totalMinutes % 60;

    const d = new Date();
    d.setHours(hh, mm, 0, 0);
    return d.getTime();
  }

  const s = String(v).trim();
  if (!s) return null;

  const hm = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (hm) {
    const hh = parseInt(hm[1], 10);
    const mm = parseInt(hm[2], 10);
    const d = new Date();
    d.setHours(hh, mm, 0, 0);
    return d.getTime();
  }

  const ampm = /^(\d{1,2}):(\d{2})\s*(am|pm)$/i.exec(s);
  if (ampm) {
    let hh = parseInt(ampm[1], 10);
    const mm = parseInt(ampm[2], 10);
    const ap = ampm[3].toLowerCase();

    if (ap === "pm" && hh !== 12) hh += 12;
    if (ap === "am" && hh === 12) hh = 0;

    const d = new Date();
    d.setHours(hh, mm, 0, 0);
    return d.getTime();
  }

  return null;
}

/* =========================
   Table rendering
   ========================= */
function resetTableState() {
  displayRowsCurrent = [];
  currentPage = 1;
  sortCol = 2;
  sortDir = "asc";

  if (els.thead) els.thead.innerHTML = "";
  if (els.tbody) els.tbody.innerHTML = "";
  if (els.pagination) els.pagination.innerHTML = "";
  if (els.emptyState) els.emptyState.style.display = "none";
}

function setStatusLabel(label, count) {
  if (!els.status) return;
  els.status.innerHTML = `<strong>${escapeHtml(`${label}: ${count} “BOOKED”`)}</strong>`;
}

function drawHeader(headers) {
  if (!els.thead) return;
  els.thead.innerHTML = "";

  const tr = document.createElement("tr");

  headers.forEach((h, idx) => {
    let label = String(h);

    if ((label === "Start Time" || label === "End Time") && tzInfo?.abbr) {
      label = `${label} (${tzInfo.abbr})`;
    }

    const th = document.createElement("th");
    th.dataset.col = String(idx);
    th.innerHTML = `<span class="sort">${escapeHtml(label)} <span class="arrows"></span></span>`;
    th.addEventListener("click", () => onHeaderClick(idx));
    tr.appendChild(th);
  });

  els.thead.appendChild(tr);
  updateHeaderState();
}

function onHeaderClick(idx) {
  if (sortCol === idx) sortDir = sortDir === "asc" ? "desc" : "asc";
  else {
    sortCol = idx;
    sortDir = "asc";
  }
  currentPage = 1;
  renderTable();
  updateHeaderState();
}

function updateHeaderState() {
  els.thead?.querySelectorAll("th").forEach((th) => {
    th.classList.remove("active", "asc", "desc");
    const idx = parseInt(th.dataset.col, 10);
    if (idx === sortCol) th.classList.add("active", sortDir);
  });
}

function renderTable() {
  const rows = displayRowsCurrent.slice();

  rows.sort((a, b) => {
    const va = a?.[sortCol] ?? "";
    const vb = b?.[sortCol] ?? "";

    if (sortCol === 2) {
      const ta = Date.parse(String(va)) || -Infinity;
      const tb = Date.parse(String(vb)) || -Infinity;
      return sortDir === "asc" ? ta - tb : tb - ta;
    }

    const cmp = String(va).toLowerCase() < String(vb).toLowerCase()
      ? -1
      : String(va).toLowerCase() > String(vb).toLowerCase()
      ? 1
      : 0;

    return sortDir === "asc" ? cmp : -cmp;
  });

  const total = rows.length;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  currentPage = Math.min(Math.max(1, currentPage), totalPages);

  const start = (currentPage - 1) * pageSize;
  const pageRows = rows.slice(start, start + pageSize);

  els.tbody.innerHTML = "";

  if (!pageRows.length) {
    els.emptyState && (els.emptyState.style.display = "block");
  } else {
    els.emptyState && (els.emptyState.style.display = "none");

    const frag = document.createDocumentFragment();

    pageRows.forEach((row) => {
      const tr = document.createElement("tr");
      row.forEach((cell) => {
        const td = document.createElement("td");
        td.textContent = cell == null ? "" : String(cell);
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    });

    els.tbody.appendChild(frag);
  }

  drawPagination(totalPages);
}

function drawPagination(totalPages) {
  const nav = els.pagination;
  nav.innerHTML = "";

  const mk = (label, disabled, handler, cls = "page-btn") => {
    const b = document.createElement("button");
    b.textContent = label;
    b.className = cls;
    if (disabled) b.setAttribute("disabled", "disabled");
    b.addEventListener("click", handler);
    return b;
  };

  nav.appendChild(mk("First", currentPage === 1, () => { currentPage = 1; renderTable(); }));
  nav.appendChild(mk("Prev", currentPage === 1, () => { currentPage = Math.max(1, currentPage - 1); renderTable(); }));

  const windowSize = 5;
  let start = Math.max(1, currentPage - Math.floor(windowSize / 2));
  let end = Math.min(totalPages, start + windowSize - 1);
  start = Math.max(1, end - windowSize + 1);

  for (let p = start; p <= end; p++) {
    const b = document.createElement("button");
    b.textContent = String(p);
    b.className = "page-num" + (p === currentPage ? " active" : "");
    b.addEventListener("click", () => { currentPage = p; renderTable(); });
    nav.appendChild(b);
  }

  nav.appendChild(mk("Next", currentPage === totalPages, () => { currentPage = Math.min(totalPages, currentPage + 1); renderTable(); }));
  nav.appendChild(mk("Last", currentPage === totalPages, () => { currentPage = totalPages; renderTable(); }));
}

/* =========================
   Utils
   ========================= */
function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (s) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[s])
  );
}

function fmtLocalDate(ms) {
  const d = new Date(ms);
  return new Intl.DateTimeFormat(undefined, { year: "numeric", month: "numeric", day: "numeric" }).format(d);
}

function fmtLocalTime(ms) {
  const d = new Date(ms);
  return new Intl.DateTimeFormat(undefined, { hour: "numeric", minute: "2-digit" }).format(d);
}

/* =========================
   Timezone chip
   ========================= */
function gmtOffsetAbbr(d = new Date()) {
  const offMin = -d.getTimezoneOffset();
  const sign = offMin >= 0 ? "+" : "-";
  const abs = Math.abs(offMin);
  const hh = Math.floor(abs / 60);
  const mm = abs % 60;
  return "GMT" + sign + String(hh) + (mm ? ":" + String(mm).padStart(2, "0") : "");
}

function detectTimezone() {
  const iana = (Intl.DateTimeFormat().resolvedOptions().timeZone || "").trim();
  const parts = new Intl.DateTimeFormat(undefined, { timeZoneName: "short" }).formatToParts(new Date());
  const tzPart = parts.find((p) => p.type === "timeZoneName");

  let abbr = tzPart && tzPart.value ? tzPart.value.trim() : "";
  const looksAlpha = /^[A-Za-z]{2,5}$/.test(abbr);
  const looksGMT = /^GMT[+-]/i.test(abbr);

  if (!abbr || (!looksAlpha && !looksGMT)) abbr = gmtOffsetAbbr();
  return { abbr, iana };
}

function tzCombinedLabel(info = tzInfo) {
  if (!info) return "";
  if (info.abbr && info.iana) return `${info.abbr} (${info.iana})`;
  return info.abbr || info.iana || "";
}

function ensureTzChip() {
  const bar = document.querySelector(".controls");
  if (!bar) return;

  let chip = document.getElementById("tzChip");
  const label = tzCombinedLabel();
  if (!label) return;

  if (!chip) {
    chip = document.createElement("span");
    chip.id = "tzChip";
    chip.setAttribute("aria-label", "Displayed time zone");
    chip.style.marginLeft = "auto";
    chip.style.background = "#E6F6FF";
    chip.style.border = "1px solid #cbd5e1";
    chip.style.color = "#0f172a";
    chip.style.fontSize = "12px";
    chip.style.fontWeight = "700";
    chip.style.padding = "4px 8px";
    chip.style.borderRadius = "999px";
    bar.appendChild(chip);
  }

  chip.textContent = `Local time: ${label}`;
}
