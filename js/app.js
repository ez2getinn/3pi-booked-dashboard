/* ============================================================================
 * 3PI BOOKED DASHBOARD (Azure Static Web Apps)
 *
 * API Endpoints:
 *   ✅ /api/getSheetNames
 *   ✅ /api/getSheetData?sheet=NAME
 *   ✅ /api/getBookedCounts
 *
 * Table Output Columns (EXACT ORDER):
 *   Email | Name | Date | Start Time | End Time | BOOKED | Site | Account | Ticket | Shift4 MID
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

      // If MONTHLY collapses, go back to default month (Jan)
      if (id === "panel-monthly" && !nowOpen) {
        forceDefaultMonth();
        queueTask(() => reloadSelectedMonth());
      }
    });
  });
}

/* =========================
   Loader
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
  }, 100);
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

let monthTabs = [];
let activeMonthName = null;

// pagination/sort
let pageSize = (els.pageSize && parseInt(els.pageSize.value, 10)) || 15;
let currentPage = 1;
let sortCol = 2; // default sort by Date
let sortDir = "asc";

// timezone chip
let tzInfo = { abbr: "", iana: "" };

/* =========================
   Column mapping (from Excel sheet)
   Your sheet (from screenshot):
   A: Date
   B: Title
   C: Description
   D: Email
   E: Work (B) => Name
   F: Start Time
   G: End Time
   H: BOOKED
   I: Site
   J: Account
   K: Ticket
   L: Shift4 MID
   ========================= */
const COL_DATE = 0;
const COL_TITLE = 1;
const COL_DESC = 2;
const COL_EMAIL = 3;
const COL_NAME = 4;       // Work (B)
const COL_START = 5;
const COL_END = 6;
const COL_BOOKED = 7;
const COL_SITE = 8;
const COL_ACCOUNT = 9;
const COL_TICKET = 10;
const COL_MID = 11;

// output columns required (EXACT ORDER)
const OUT_COLS = [
  { key: "Email",     src: COL_EMAIL },
  { key: "Name",      src: COL_NAME },
  { key: "Date",      src: COL_DATE },
  { key: "Start Time",src: COL_START },
  { key: "End Time",  src: COL_END },
  { key: "BOOKED",    src: COL_BOOKED },
  { key: "Site",      src: COL_SITE },
  { key: "Account",   src: COL_ACCOUNT },
  { key: "Ticket",    src: COL_TICKET },
  { key: "Shift4 MID",src: COL_MID },
];

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
      await reloadSelectedMonth();
    });
  });

  els.pageSize?.addEventListener("change", () => {
    pageSize = parseInt(els.pageSize.value, 10) || 15;
    currentPage = 1;
    renderTable();
  });

  // clicking scorecard loads that month
  els.cards?.addEventListener("click", (e) => {
    const card = e.target.closest(".card[data-name]");
    if (!card) return;

    const name = card.getAttribute("data-name");
    if (!monthTabs.includes(name)) return;

    activeMonthName = name;
    setActiveCard(name);
    queueTask(() => reloadSelectedMonth());
  });

  // first load
  queueTask(() => loadTabsAndFirstPaint());
});

/* =========================
   Load tabs + default month
   ========================= */
async function loadTabsAndFirstPaint() {
  const names = await apiGet("/api/getSheetNames", "getSheetNames").catch(() => []);
  monthTabs = Array.isArray(names) ? names.slice() : [];

  if (!monthTabs.length) {
    setStatusLabel("No months found", 0);
    return;
  }

  // Default to Jan ALWAYS (if exists)
  activeMonthName = monthTabs.includes("Jan") ? "Jan" : monthTabs[0];

  await refreshCards();
  setActiveCard(activeMonthName);
  await reloadSelectedMonth();
}

function forceDefaultMonth() {
  activeMonthName = monthTabs.includes("Jan") ? "Jan" : monthTabs[0];
  setActiveCard(activeMonthName);
}

/* =========================
   Scorecards
   ========================= */
let lastCardsJSON = "";

async function refreshCards() {
  const items = await apiGet("/api/getBookedCounts", "getBookedCounts").catch((err) => {
    console.error("getBookedCounts failed:", err);
    els.cards.innerHTML =
      `<div style="padding:10px;color:#b91c1c;font-weight:900;">Failed to load monthly scorecards</div>`;
    return null;
  });

  if (!items) return;

  const j = JSON.stringify(items);
  if (j === lastCardsJSON) return;

  lastCardsJSON = j;
  renderCards(items);
  setActiveCard(activeMonthName);
}

function renderCards(items) {
  if (!Array.isArray(items)) {
    els.cards.innerHTML = "";
    return;
  }

  els.cards.innerHTML = items
    .map(({ name, count }) => {
      return `
        <div class="card" data-name="${escapeHtml(name)}" title="${escapeHtml(name)}">
          <span class="chip">BOOKED</span>
          <div class="title">${escapeHtml(name)}</div>
          <div class="value">${Number(count || 0)}</div>
        </div>
      `;
    })
    .join("");
}

function setActiveCard(name) {
  els.cards?.querySelectorAll(".card").forEach((c) => {
    c.classList.toggle("active", c.getAttribute("data-name") === name);
  });
}

/* =========================
   Reload month sheet
   ========================= */
let rawRows = [];

async function reloadSelectedMonth() {
  if (!activeMonthName) return;

  resetTableState();

  const payload = await apiGet(
    `/api/getSheetData?sheet=${encodeURIComponent(activeMonthName)}`,
    `getSheetData:${activeMonthName}`
  ).catch((err) => {
    console.error("getSheetData failed:", err);
    return { headers: [], rows: [] };
  });

  rawRows = Array.isArray(payload.rows) ? payload.rows : [];

  // Build display headers + display rows
  const displayHeaders = OUT_COLS.map((c) => c.key);

  const displayRows = rawRows.map((row) => {
    return OUT_COLS.map((c) => {
      const v = row?.[c.src];

      // Date formatting
      if (c.key === "Date") return normalizeDateCell(v);

      // Start/End formatting
      if (c.key === "Start Time") return normalizeTimeCell(v);
      if (c.key === "End Time") return normalizeTimeCell(v);

      return v == null ? "" : String(v);
    });
  });

  setStatusLabel(activeMonthName, displayRows.length);
  drawHeader(displayHeaders);
  renderTable(displayRows);
}

/* =========================
   Table rendering
   ========================= */
let displayRowsCurrent = [];

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
  els.status.innerHTML = `<strong>${escapeHtml(`${label}: ${count} “BOOKED”`)}</strong>`;
}

function drawHeader(headers) {
  if (!els.thead) return;
  els.thead.innerHTML = "";

  const tr = document.createElement("tr");

  headers.forEach((h, idx) => {
    let label = String(h);

    // Append timezone to Start/End headers
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

function renderTable(incomingRows) {
  if (Array.isArray(incomingRows)) displayRowsCurrent = incomingRows;

  const rows = displayRowsCurrent.slice();

  // Sort: special case Date column (index 2)
  rows.sort((a, b) => {
    const va = a?.[sortCol] ?? "";
    const vb = b?.[sortCol] ?? "";

    // Date sort (col 2)
    if (sortCol === 2) {
      const ta = parseDateToMs(va);
      const tb = parseDateToMs(vb);
      return sortDir === "asc" ? ta - tb : tb - ta;
    }

    // Start time (col 3) / End time (col 4) sort
    if (sortCol === 3 || sortCol === 4) {
      const ta = parseTimeToMinutes(va);
      const tb = parseTimeToMinutes(vb);
      return sortDir === "asc" ? ta - tb : tb - ta;
    }

    const na = Number(va);
    const nb = Number(vb);

    const bothNum =
      !Number.isNaN(na) &&
      !Number.isNaN(nb) &&
      String(va).trim() !== "" &&
      String(vb).trim() !== "";

    const cmp = bothNum
      ? na - nb
      : String(va).toLowerCase() < String(vb).toLowerCase()
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

  nav.appendChild(
    mk("First", currentPage === 1, () => {
      currentPage = 1;
      renderTable();
    })
  );
  nav.appendChild(
    mk("Prev", currentPage === 1, () => {
      currentPage = Math.max(1, currentPage - 1);
      renderTable();
    })
  );

  const windowSize = 5;
  let start = Math.max(1, currentPage - Math.floor(windowSize / 2));
  let end = Math.min(totalPages, start + windowSize - 1);
  start = Math.max(1, end - windowSize + 1);

  for (let p = start; p <= end; p++) {
    const b = document.createElement("button");
    b.textContent = String(p);
    b.className = "page-num" + (p === currentPage ? " active" : "");
    b.addEventListener("click", () => {
      currentPage = p;
      renderTable();
    });
    nav.appendChild(b);
  }

  nav.appendChild(
    mk("Next", currentPage === totalPages, () => {
      currentPage = Math.min(totalPages, currentPage + 1);
      renderTable();
    })
  );
  nav.appendChild(
    mk("Last", currentPage === totalPages, () => {
      currentPage = totalPages;
      renderTable();
    })
  );
}

/* =========================
   Data formatting helpers
   ========================= */
function normalizeDateCell(v) {
  if (v == null) return "";

  // If Excel gave a number, it might be an Excel serial date
  if (typeof v === "number" && Number.isFinite(v)) {
    // Excel serial date: day 1 = 1899-12-31 (with 1900 leap bug)
    // This formula works for most modern Excel sheets.
    const ms = Math.round((v - 25569) * 86400 * 1000);
    return fmtLocalDate(ms);
  }

  const s = String(v).trim();
  if (!s) return "";

  // If it's already a date-like string, try Date parse
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) return fmtLocalDate(d.getTime());

  return s;
}

function normalizeTimeCell(v) {
  if (v == null) return "";

  // If Excel gave "18:00" as string, keep it but normalize display
  const s = String(v).trim();
  if (!s) return "";

  // if it contains AM/PM already, keep it
  if (/\b(am|pm)\b/i.test(s)) return s;

  // If it's "18:00" or "18:00:00" -> show in 12-hour format
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (m) {
    const hh = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);

    const d = new Date();
    d.setHours(hh, mm, 0, 0);

    return new Intl.DateTimeFormat(undefined, {
      hour: "numeric",
      minute: "2-digit",
    }).format(d);
  }

  return s;
}

function parseDateToMs(s) {
  const t = Date.parse(String(s || "").trim());
  return Number.isNaN(t) ? -Infinity : t;
}

function parseTimeToMinutes(s) {
  const str = String(s || "").trim().toLowerCase();
  if (!str) return -Infinity;

  // 9:30 AM
  const ampm = /^(\d{1,2}):(\d{2})\s*(am|pm)$/.exec(str);
  if (ampm) {
    let hh = parseInt(ampm[1], 10);
    const mm = parseInt(ampm[2], 10);
    const ap = ampm[3];

    if (ap === "pm" && hh !== 12) hh += 12;
    if (ap === "am" && hh === 12) hh = 0;

    return hh * 60 + mm;
  }

  // 18:00
  const hm = /^(\d{1,2}):(\d{2})$/.exec(str);
  if (hm) return parseInt(hm[1], 10) * 60 + parseInt(hm[2], 10);

  return -Infinity;
}

/* =========================
   Utils
   ========================= */
function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (s) =>
    ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    }[s])
  );
}

function fmtLocalDate(ms) {
  if (ms == null) return "";
  const d = new Date(ms);
  return new Intl.DateTimeFormat(undefined, {
    year: "numeric",
    month: "numeric",
    day: "numeric",
  }).format(d);
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
  const parts = new Intl.DateTimeFormat(undefined, {
    timeZoneName: "short",
  }).formatToParts(new Date());
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
