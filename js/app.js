/* ============================================================================
 * SHIFT4 Booked Dashboard – app.js (Azure Static Web Apps + Azure Functions)
 *
 * ✅ Frontend calls Azure Functions:
 *   GET /api/getSheetNames
 *   GET /api/getSheetData?sheet=Sep%202025
 *
 * ✅ Scorecards:
 *   - Calculated in browser (no need /api/getBookedCounts)
 *
 * ✅ No Google Apps Script logic (REMOVED)
 * ✅ No /api/getVersion polling (REMOVED)
 * ========================================================================== */

/* ========== Azure Functions fetch wrapper ========== */
async function apiGet(path, label = path) {
  const url = String(path || "").trim();

  try {
    const res = await fetch(url, {
      method: "GET",
      headers: { Accept: "application/json" },
    });

    const ct = (res.headers.get("content-type") || "").toLowerCase();
    const text = await res.text();

    if (!res.ok) {
      throw new Error(`[${label}] HTTP ${res.status}: ${text.slice(0, 200)}`);
    }

    // If Azure returns HTML, API route isn't being served (or wrong URL)
    if (ct.includes("text/html")) {
      throw new Error(
        `[${label}] Returned HTML instead of JSON. API not found or misrouted.`
      );
    }

    try {
      return JSON.parse(text);
    } catch {
      throw new Error(`[${label}] Invalid JSON: ${text.slice(0, 200)}`);
    }
  } catch (err) {
    console.warn(`apiGet failed: ${label}`, err);
    throw err;
  }
}

/* ========== Collapsibles ========== */
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

      // MONTHLY closed or YEARLY opened → Default view (Today→Dec)
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

/* ========== Loader ========== */
let pendingLoads = 0;
function showLoader() {
  pendingLoads++;
  document.getElementById("loaderOverlay")?.classList.remove("hidden");
}
function hideLoader() {
  pendingLoads = Math.max(0, pendingLoads - 1);
  if (!pendingLoads)
    document.getElementById("loaderOverlay")?.classList.add("hidden");
}

/* ========== Gate / Debounce (avoid bursts) ========== */
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

/* ========== DOM & App State ========== */
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

// Modes
let defaultMode = true;
let fromTodayLowerBound = true;
let toEndOfYearUpperBound = null;

// Table state
let currentRows = [];
let currentRowsMs = [];
let currentHeaders = [];
let currentPage = 1;
let pageSize = (els.pageSize && parseInt(els.pageSize.value, 10)) || 15;
let sortCol = 2;
let sortDir = "asc";

const DATE_COL_INDEX = 2;
const START_COL_INDEX = 3;
const END_COL_INDEX = 4;

// Timezone display state
let tzInfo = { abbr: "", iana: "" };

/* ========== Month parsing / dates ========== */
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

const MONTHS_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

const MONTHS_LONG = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

function parseTab(name) {
  const m = MONTH_RX.exec(String(name).trim());
  if (!m) return null;
  const key = m[1].slice(0, 3).toLowerCase();
  return {
    month: MONTH_MAP[key] ?? MONTH_MAP[m[1].toLowerCase()],
    year: m[2] ? parseInt(m[2], 10) : null,
  };
}

function endOfYear(d) {
  return new Date(d.getFullYear(), 11, 31, 23, 59, 59, 999);
}

function pickCurrentMonthIndex(names) {
  const now = new Date(),
    y = now.getFullYear(),
    m = now.getMonth();

  const short = MONTHS_SHORT[m],
    long = MONTHS_LONG[m];

  let rx = new RegExp(`^(?:${short}|${long})\\s+${y}$`, "i");
  let i = names.findIndex((n) => rx.test(n));
  if (i >= 0) return i;

  rx = new RegExp(`^(?:${short}|${long})$`, "i");
  i = names.findIndex((n) => rx.test(n));
  return i >= 0 ? i : 0;
}

/* ========== Boot ========== */
document.addEventListener("DOMContentLoaded", () => {
  initCollapsibles();

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

  // Scorecard click → single month view
  els.cards?.addEventListener("click", (e) => {
    const card = e.target.closest(".card[data-name]");
    if (!card) return;

    const name = card.getAttribute("data-name");
    if (!monthTabs.includes(name)) return;

    activeMonthName = name;
    defaultMode = false;
    fromTodayLowerBound = false;
    toEndOfYearUpperBound = null;

    setActiveCard(name);
    queueTask(() => reload());
  });

  // first paint
  queueTask(() => loadTabsAndFirstPaint());
});

/* ========== First paint ========== */
async function loadTabsAndFirstPaint() {
  tzInfo = detectTimezone();
  ensureTzChip();

  const names = await apiGet("/api/getSheetNames", "getSheetNames").catch(
    () => []
  );
  monthTabs = Array.isArray(names) ? names.slice() : [];

  if (!monthTabs.length) {
    setSummaryLabel(false, "No month", 0);
    els.cards && (els.cards.innerHTML = "");
    return;
  }

  selectCurrentMonth();

  // scorecards
  await refreshCards();

  // default mode (Today → Dec)
  forceDefaultMode();
  await reload();

  // optional heartbeat refresh (safe)
  startCardsHeartbeat();
}

function selectCurrentMonth() {
  const idx = pickCurrentMonthIndex(monthTabs);
  activeMonthName = monthTabs[Math.max(0, idx)];
  setActiveCard(activeMonthName);
}

/* ========== Scorecards (BOOKED counts) ========== */
let lastCardsJSON = "";

async function refreshCards() {
  // Fallback-only version: calculate counts from each month sheet
  if (!Array.isArray(monthTabs) || !monthTabs.length) {
    els.cards && (els.cards.innerHTML = "");
    return;
  }

  const results = [];

  for (const name of monthTabs) {
    const payload = await apiGet(
      `/api/getSheetData?sheet=${encodeURIComponent(name)}`,
      `getSheetData:${name}`
    ).catch(() => null);

    if (!payload || !Array.isArray(payload.headers) || !Array.isArray(payload.rows)) {
      results.push({ name, count: 0 });
      continue;
    }

    // Find "Booked" column index
    const bookedIdx = payload.headers.findIndex(
      (h) => String(h || "").trim().toLowerCase() === "booked"
    );

    let count = 0;
    if (bookedIdx >= 0) {
      for (const r of payload.rows) {
        const v = String((r && r[bookedIdx]) || "").trim().toLowerCase();
        if (v === "booked") count++;
      }
    }

    results.push({ name, count });
  }

  const j = JSON.stringify(results);
  if (j === lastCardsJSON) return;

  lastCardsJSON = j;
  renderCards(results);
  setActiveCard(activeMonthName);
}

function renderCards(items) {
  if (!Array.isArray(items)) {
    els.cards.innerHTML = "";
    return;
  }

  els.cards.innerHTML = items
    .map(
      ({ name, count }) => `
      <div class="card" data-name="${escapeHtml(name)}" title="${escapeHtml(name)}">
        <span class="chip">BOOKED</span>
        <div class="title">${escapeHtml(name)}</div>
        <div class="value">${count}</div>
      </div>
    `
    )
    .join("");
}

function setActiveCard(name) {
  els.cards?.querySelectorAll(".card").forEach((c) => {
    c.classList.toggle("active", c.getAttribute("data-name") === name);
  });
}

/* ========== Summary label helper ========== */
function setSummaryLabel(isDefaultMode, baseLabel, count) {
  const label = `${baseLabel}: ${count} “BOOKED”`;
  els.status.innerHTML = `<strong>${escapeHtml(label)}</strong>`;
}

/* ========== Mode helpers ========== */
function forceDefaultMode() {
  defaultMode = true;
  selectCurrentMonth();
  fromTodayLowerBound = true;
  toEndOfYearUpperBound = endOfYear(new Date());
}

function resetViewState() {
  els.thead && (els.thead.innerHTML = "");
  els.tbody && (els.tbody.innerHTML = "");
  els.emptyState && (els.emptyState.style.display = "none");
  els.pagination && (els.pagination.innerHTML = "");

  currentRows = [];
  currentRowsMs = [];
  currentPage = 1;
  sortCol = DATE_COL_INDEX;
  sortDir = "asc";
}

/* ========== Reload paths ========== */
async function reload() {
  const monthlyOpen = document
    .getElementById("panel-monthly")
    ?.classList.contains("open");

  if (!monthlyOpen) forceDefaultMode();
  return defaultMode ? reloadDefaultRange() : reloadSingleMonth(activeMonthName);
}

// Single month view
async function reloadSingleMonth(tab) {
  if (!tab) return;
  resetViewState();

  const payload = await apiGet(
    `/api/getSheetData?sheet=${encodeURIComponent(tab)}`,
    `getSheetData:${tab}`
  ).catch(() => ({ headers: [], rows: [], ms: [] }));

  currentHeaders = payload.headers || [];
  currentRows = Array.isArray(payload.rows) ? payload.rows : [];
  currentRowsMs = Array.isArray(payload.ms) ? payload.ms : [];

  drawHeader(currentHeaders);
  renderTable();

  setSummaryLabel(false, tab, currentRows.length);
  setActiveCard(tab);
}

// Default range view (today→Dec 31)
async function reloadDefaultRange() {
  if (!monthTabs.length) return;
  resetViewState();

  const now = new Date(),
    y = now.getFullYear(),
    m = now.getMonth();

  const wanted = monthTabs.filter((n) => {
    const p = parseTab(n);
    if (!p) return false;
    const year = p.year == null ? y : p.year;
    return year === y && p.month >= m;
  });

  const list = wanted.length
    ? wanted
    : monthTabs.slice(pickCurrentMonthIndex(monthTabs));

  let headersPicked = false;
  currentRows = [];
  currentRowsMs = [];

  for (const name of list) {
    const d = await apiGet(
      `/api/getSheetData?sheet=${encodeURIComponent(name)}`,
      `getSheetData:${name}`
    ).catch(() => ({ headers: [], rows: [], ms: [] }));

    if (!headersPicked && d.headers && d.headers.length) {
      currentHeaders = d.headers;
      headersPicked = true;
    }

    if (Array.isArray(d.rows)) currentRows.push(...d.rows);
    if (Array.isArray(d.ms)) currentRowsMs.push(...d.ms);
  }

  fromTodayLowerBound = true;
  toEndOfYearUpperBound = endOfYear(now);

  drawHeader(currentHeaders);
  renderTable();

  const label = `${MONTHS_SHORT[m]}–Dec`;
  const totalInWindow = applyDateWindow(currentRows, currentRowsMs).length;
  setSummaryLabel(true, label, totalInWindow);
}

/* ========== Table rendering (sort/paginate) ========== */
function drawHeader(headers) {
  if (!els.thead) return;
  els.thead.innerHTML = "";
  const tr = document.createElement("tr");

  const shortTz = tzInfo?.abbr || "";

  (headers || []).forEach((h, idx) => {
    let label = String(h ?? "");

    if ((idx === START_COL_INDEX || idx === END_COL_INDEX) && shortTz) {
      if (
        !/\(\s*[A-Za-z]{2,5}\s*\)$/i.test(label) &&
        !/\(GMT[+-]\d{1,2}(?::\d{2})?\)$/i.test(label)
      ) {
        label += ` (${shortTz})`;
      }
    }

    const th = document.createElement("th");
    th.dataset.col = String(idx);
    th.innerHTML = `<span class="sort">${escapeHtml(
      label
    )} <span class="arrows"></span></span>`;
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

function applyDateWindow(rows, rowsMs) {
  if (
    !Array.isArray(rows) ||
    !Array.isArray(rowsMs) ||
    rows.length !== rowsMs.length
  )
    return [];

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const t0 = fromTodayLowerBound ? today.getTime() : -Infinity;
  const t1 = toEndOfYearUpperBound
    ? toEndOfYearUpperBound.getTime()
    : Infinity;

  const items = [];
  for (let i = 0; i < rows.length; i++) {
    const ms = rowsMs[i]?.dateMs ?? null;
    if (ms == null) continue;
    if (ms >= t0 && ms <= t1) items.push({ row: rows[i], ms: rowsMs[i], idx: i });
  }
  return items;
}

function renderTable() {
  let items;
  if (defaultMode) {
    items = applyDateWindow(currentRows, currentRowsMs);
  } else {
    items = currentRows.map((row, i) => ({
      row,
      ms: currentRowsMs[i] || { dateMs: null, startMs: null, endMs: null },
      idx: i,
    }));
  }

  items.sort((a, b) => {
    const va = a.row[sortCol];
    const vb = b.row[sortCol];

    if (sortCol === DATE_COL_INDEX) {
      const ta = a.ms?.dateMs ?? -Infinity;
      const tb = b.ms?.dateMs ?? -Infinity;
      return sortDir === "asc" ? ta - tb : tb - ta;
    }

    if (sortCol === START_COL_INDEX || sortCol === END_COL_INDEX) {
      const key = sortCol === START_COL_INDEX ? "startMs" : "endMs";
      const ta = a.ms?.[key] ?? -Infinity;
      const tb = b.ms?.[key] ?? -Infinity;
      if (ta !== tb) return sortDir === "asc" ? ta - tb : tb - ta;

      const da = a.ms?.dateMs ?? -Infinity;
      const db = b.ms?.dateMs ?? -Infinity;
      return sortDir === "asc" ? da - db : db - da;
    }

    const na = Number(va);
    const nb = Number(vb);

    const both =
      !Number.isNaN(na) &&
      !Number.isNaN(nb) &&
      String(va).trim() !== "" &&
      String(vb).trim() !== "";

    const cmp = both
      ? na - nb
      : String(va).toLowerCase() < String(vb).toLowerCase()
      ? -1
      : String(va).toLowerCase() > String(vb).toLowerCase()
      ? 1
      : 0;

    return sortDir === "asc" ? cmp : -cmp;
  });

  const total = items.length;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  currentPage = Math.min(Math.max(1, currentPage), totalPages);

  const start = (currentPage - 1) * pageSize;
  const pageItems = items.slice(start, start + pageSize);

  els.tbody.innerHTML = "";

  if (!pageItems.length) {
    els.emptyState && (els.emptyState.style.display = "block");
  } else {
    els.emptyState && (els.emptyState.style.display = "none");
    const frag = document.createDocumentFragment();

    pageItems.forEach(({ row, ms }) => {
      const tr = document.createElement("tr");
      const cells = row.slice();

      cells[DATE_COL_INDEX] = fmtLocalDate(ms?.dateMs ?? null);
      cells[START_COL_INDEX] = fmtLocalTime(ms?.startMs ?? null);
      cells[END_COL_INDEX] = fmtLocalTime(ms?.endMs ?? null);

      cells.forEach((cell) => {
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

/* ========== Card heartbeat ========== */
const CARDS_HEARTBEAT_MS = 120 * 1000;
let cardsTimer = null;

function startCardsHeartbeat() {
  cardsTimer = setInterval(() => {
    if (netBusy) return;
    queueTask(async () => {
      const prev = lastCardsJSON;
      await refreshCards();
      if (lastCardsJSON !== prev) {
        forceDefaultMode();
        await reload();
      }
    });
  }, CARDS_HEARTBEAT_MS);
}

/* ========== Utils ========== */
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

function fmtLocalTime(ms) {
  if (ms == null) return "";
  const d = new Date(ms);
  return new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(d);
}

/* ========== Timezone helpers ========== */
function gmtOffsetAbbr(d = new Date()) {
  const offMin = -d.getTimezoneOffset();
  const sign = offMin >= 0 ? "+" : "-";
  const abs = Math.abs(offMin);
  const hh = Math.floor(abs / 60);
  const mm = abs % 60;
  return (
    "GMT" +
    sign +
    String(hh) +
    (mm ? ":" + String(mm).padStart(2, "0") : "")
  );
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
