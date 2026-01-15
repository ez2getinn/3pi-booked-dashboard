/* ============================================================================
 * 3PI BOOKED DASHBOARD – FRONTEND (Azure Static Web Apps)
 *
 * DEFAULT MODE (Option B):
 *   ✅ Today → End of Year (across remaining months)
 *
 * SCORECARD MODE:
 *   ✅ Click a month → Full month view
 *
 * APIs REQUIRED:
 *   ✅ /api/getSheetNames
 *   ✅ /api/getSheetData?sheet=NAME
 *   ✅ /api/getBookedCounts
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
    cache: "no-store",
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

      // Monthly closed OR Yearly opened -> Default Mode reload
      if (
        (id === "panel-monthly" && !nowOpen) ||
        (id === "panel-yearly" && nowOpen)
      ) {
        forceDefaultMode();
        queueTask(() => reload());
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
   Excel column mapping (your sheet)
   =========================
   A=Date
   B=Title (ignore)
   C=Description (ignore)
   D=Email
   E=Work (B) => Name
   F=Start Time
   G=End Time
   H=BOOKED
   I=Site
   J=Account
   K=Ticket
   L=Shift4 MID
*/
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

/* Required output columns (EXACT ORDER) */
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
   Month detection + ordering
   ========================= */
const MONTH_RX =
  /^(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)(?:\s+(\d{4}))?$/i;

const MONTH_MAP = {
  jan: 0,
  feb: 1,
  mar: 2,
  apr: 3,
  may: 4,
  jun: 5,
  jul: 6,
  aug: 7,
  sep: 8,
  sept: 8,
  oct: 9,
  nov: 10,
  dec: 11,
};

const MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function parseMonthSheet(name) {
  const clean = String(name || "").trim();
  if (!clean) return null;

  const m = MONTH_RX.exec(clean);
  if (!m) return null;

  const key = m[1].slice(0, 3).toLowerCase();
  const monthIndex = MONTH_MAP[key];
  const year = m[2] ? parseInt(m[2], 10) : null;

  if (monthIndex == null) return null;
  return { monthIndex, year };
}

function isMonthSheet(name) {
  return !!parseMonthSheet(name);
}

function sortMonthNamesJanToDec(list) {
  return list.slice().sort((a, b) => {
    const pa = parseMonthSheet(a);
    const pb = parseMonthSheet(b);

    const ya = pa?.year ?? 9999;
    const yb = pb?.year ?? 9999;
    if (ya !== yb) return ya - yb;

    return (pa?.monthIndex ?? 99) - (pb?.monthIndex ?? 99);
  });
}

function pickCurrentMonthTab(names) {
  const now = new Date();
  const m = now.getMonth();
  const short = MONTHS_SHORT[m].toLowerCase();

  const found = names.find((n) => String(n).trim().slice(0, 3).toLowerCase() === short);
  return found || names[0];
}

/* =========================
   Timezone chip (viewer local)
   ========================= */
let tzInfo = { abbr: "", iana: "" };

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

/* =========================
   App state
   ========================= */
let monthTabs = [];
let activeMonthName = null;

// defaultMode true = Today->EndOfYear across remaining months
let defaultMode = true;

let displayRowsCurrent = [];
let currentPage = 1;
let pageSize = (els.pageSize && parseInt(els.pageSize.value, 10)) || 15;

let sortCol = 2; // Date index in output table
let sortDir = "asc";

let filterStartMs = null;
let filterEndMs = null;

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
      await reload();
    });
  });

  els.pageSize?.addEventListener("change", () => {
    pageSize = parseInt(els.pageSize.value, 10) || 15;
    currentPage = 1;
    renderTable();
  });

  // Scorecard click = Full month view
  els.cards?.addEventListener("click", (e) => {
    const card = e.target.closest(".card[data-name]");
    if (!card) return;

    const name = card.getAttribute("data-name");
    if (!monthTabs.includes(name)) return;

    activeMonthName = name;
    defaultMode = false;

    setActiveCard(activeMonthName);
    queueTask(() => reload());
  });

  queueTask(() => loadTabsAndFirstPaint());
});

/* =========================
   First paint
   ========================= */
async function loadTabsAndFirstPaint() {
  // ✅ Get sheet tabs (already filtered on backend)
  const names = await apiGet("/api/getSheetNames", "getSheetNames").catch(() => []);
  monthTabs = Array.isArray(names) ? names.filter(isMonthSheet) : [];
  monthTabs = sortMonthNamesJanToDec(monthTabs);

  if (!monthTabs.length) {
    setStatusLabel("No month sheets found", 0);
    return;
  }

  // ✅ Active month is current month
  activeMonthName = pickCurrentMonthTab(monthTabs);

  // ✅ Default mode
  forceDefaultMode();

  // ✅ Scorecards
  await refreshCards();

  // ✅ Load table
  await reload();
}

function endOfYear(d) {
  return new Date(d.getFullYear(), 11, 31, 23, 59, 59, 999);
}

function forceDefaultMode() {
  defaultMode = true;

  // Start = today 00:00 local
  const now = new Date();
  const start = new Date(now);
  start.setHours(0, 0, 0, 0);

  filterStartMs = start.getTime();
  filterEndMs = endOfYear(now).getTime();

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

  // Enforce Jan->Dec on frontend too
  items = items
    .filter((x) => x && isMonthSheet(x.name))
    .sort((a, b) => (parseMonthSheet(a.name)?.monthIndex ?? 99) - (parseMonthSheet(b.name)?.monthIndex ?? 99));

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
   Reload logic
   ========================= */
async function reload() {
  const monthlyOpen = document
    .getElementById("panel-monthly")
    ?.classList.contains("open");

  if (!monthlyOpen) forceDefaultMode();

  if (defaultMode) return reloadDefaultTodayToEndOfYear();
  return reloadSingleMonth(activeMonthName);
}

/* ✅ Default mode = Today → End of Year across remaining months */
async function reloadDefaultTodayToEndOfYear() {
  resetTableState();

  const now = new Date();
  const currentMonthIndex = now.getMonth();

  // Only current month -> Dec within same year
  const list = monthTabs.filter((n) => {
    const p = parseMonthSheet(n);
    if (!p) return false;

    const year = p.year ?? now.getFullYear();
    if (year !== now.getFullYear()) return false;

    return p.monthIndex >= currentMonthIndex;
  });

  const rowsAll = [];

  for (const sheetName of list) {
    const payload = await apiGet(
      `/api/getSheetData?sheet=${encodeURIComponent(sheetName)}`,
      `getSheetData:${sheetName}`
    ).catch(() => ({ rows: [] }));

    const rawRows = Array.isArray(payload.rows) ? payload.rows : [];

    // only BOOKED rows
    for (const r of rawRows) {
      if (String(r?.[COL_BOOKED] || "").trim().toUpperCase() !== "BOOKED") continue;
      rowsAll.push(r);
    }
  }

  const mapped = mapRowsForDisplay(rowsAll);

  // Filter today->end of year by actual Date column
  const filtered = mapped.filter((r) => {
    const ms = r.__dateMs;
    if (!Number.isFinite(ms)) return false;
    return ms >= filterStartMs && ms <= filterEndMs;
  });

  displayRowsCurrent = filtered.map((x) => x.display);

  drawHeader(OUT_COLS.map((c) => c.key));

  const label = `${MONTHS_SHORT[currentMonthIndex]}–Dec`;
  setStatusLabel(`${label} (Today → End of Year)`, displayRowsCurrent.length);

  renderTable();
}

/* ✅ Single month view (full month) */
async function reloadSingleMonth(tab) {
  if (!tab) return;

  resetTableState();

  const payload = await apiGet(
    `/api/getSheetData?sheet=${encodeURIComponent(tab)}`,
    `getSheetData:${tab}`
  ).catch(() => ({ rows: [] }));

  const rawRows = Array.isArray(payload.rows) ? payload.rows : [];

  const bookedOnly = rawRows.filter(
    (r) => String(r?.[COL_BOOKED] || "").trim().toUpperCase() === "BOOKED"
  );

  const mapped = mapRowsForDisplay(bookedOnly);
  displayRowsCurrent = mapped.map((x) => x.display);

  drawHeader(OUT_COLS.map((c) => c.key));
  setStatusLabel(`${tab} (Full Month)`, displayRowsCurrent.length);

  renderTable();
}

/* =========================
   Mapping + formatting
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

  // Excel serial date -> JS epoch
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

  // Excel serial datetime -> fractional portion is time-of-day
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

  // 14:30
  const hm = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (hm) {
    const hh = parseInt(hm[1], 10);
    const mm = parseInt(hm[2], 10);
    const d = new Date();
    d.setHours(hh, mm, 0, 0);
    return d.getTime();
  }

  // 2:30 PM
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
   Table render
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

    // Date sorting
    if (sortCol === 2) {
      const ta = Date.parse(String(va)) || -Infinity;
      const tb = Date.parse(String(vb)) || -Infinity;
      return sortDir === "asc" ? ta - tb : tb - ta;
    }

    const cmp =
      String(va).toLowerCase() < String(vb).toLowerCase()
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
  return new Intl.DateTimeFormat(undefined, {
    year: "numeric",
    month: "numeric",
    day: "numeric",
  }).format(d);
}

function fmtLocalTime(ms) {
  const d = new Date(ms);
  return new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(d);
}
