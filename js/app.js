/* ============================================================================
 * 3PI BOOKED DASHBOARD (Azure Static Web Apps Frontend)
 *
 * Uses Azure Functions:
 *   ✅ /api/getSheetNames
 *   ✅ /api/getSheetData?sheet=NAME
 *   ✅ /api/getBookedCounts
 * ========================================================================== */

/* ========== Azure API fetch wrapper ========== */
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
  } catch (e) {
    console.warn(`[API:${label}] Invalid JSON:`, text.slice(0, 500));
    throw new Error(`Invalid JSON returned for: ${label}`);
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

      // If MONTHLY is closed -> go back to default (Jan)
      if (id === "panel-monthly" && !nowOpen) {
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
  if (!pendingLoads) document.getElementById("loaderOverlay")?.classList.add("hidden");
}

/* ========== Gate / Debounce ========== */
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

// Modes:
// defaultMode = true means "Yearly default view" (but we want Jan on load)
let defaultMode = true;

// Table state
let currentRows = [];
let currentHeaders = [];
let currentPage = 1;
let pageSize = (els.pageSize && parseInt(els.pageSize.value, 10)) || 15;
let sortCol = 0;
let sortDir = "asc";

// ✅ IMPORTANT: Excel columns indexes (based on your API headers)
const DATE_COL_INDEX = 0;      // Date
const START_COL_INDEX = 5;     // Start Time
const END_COL_INDEX = 6;       // End Time

// Timezone display state
let tzInfo = { abbr: "", iana: "" };

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

  // Scorecard click -> single month view
  els.cards?.addEventListener("click", (e) => {
    const card = e.target.closest(".card[data-name]");
    if (!card) return;

    const name = card.getAttribute("data-name");
    if (!monthTabs.includes(name)) return;

    activeMonthName = name;
    defaultMode = false;

    setActiveCard(name);
    queueTask(() => reloadSingleMonth(activeMonthName));
  });

  // FIRST LOAD
  queueTask(() => loadTabsAndFirstPaint());
});

/* ========== First paint ========== */
async function loadTabsAndFirstPaint() {
  tzInfo = detectTimezone();
  ensureTzChip();

  const names = await apiGet("/api/getSheetNames", "getSheetNames").catch(() => []);
  monthTabs = Array.isArray(names) ? names.slice() : [];

  if (!monthTabs.length) {
    setSummaryLabel("No months found", 0);
    return;
  }

  // ✅ Always default to Jan on load if it exists
  activeMonthName = monthTabs.includes("Jan") ? "Jan" : monthTabs[0];

  // Load Scorecards
  await refreshCards();

  // Load Jan into table on page load
  defaultMode = false;
  setActiveCard(activeMonthName);
  await reloadSingleMonth(activeMonthName);
}

/* ========== Scorecards ========== */
let lastCardsJSON = "";

async function refreshCards() {
  const items = await apiGet("/api/getBookedCounts", "getBookedCounts").catch((err) => {
    console.error("getBookedCounts failed:", err);
    els.cards.innerHTML =
      `<div style="padding:10px;color:#b91c1c;font-weight:800;">Failed to load scorecards</div>`;
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
    .map(
      ({ name, count }) => `
      <div class="card" data-name="${escapeHtml(name)}" title="${escapeHtml(name)}">
        <span class="chip">BOOKED</span>
        <div class="title">${escapeHtml(name)}</div>
        <div class="value">${Number(count || 0)}</div>
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

/* ========== Summary label ========== */
function setSummaryLabel(baseLabel, count) {
  els.status.innerHTML = `<strong>${escapeHtml(`${baseLabel}: ${count} “BOOKED”`)}</strong>`;
}

/* ========== View helpers ========== */
function resetViewState() {
  if (els.thead) els.thead.innerHTML = "";
  if (els.tbody) els.tbody.innerHTML = "";
  if (els.emptyState) els.emptyState.style.display = "none";
  if (els.pagination) els.pagination.innerHTML = "";

  currentRows = [];
  currentHeaders = [];
  currentPage = 1;

  sortCol = DATE_COL_INDEX;
  sortDir = "asc";
}

function forceDefaultMode() {
  defaultMode = false;
  activeMonthName = monthTabs.includes("Jan") ? "Jan" : monthTabs[0];
}

/* ========== Reload ========== */
async function reload() {
  // If defaultMode true, we could load all months, but user wants Jan by default.
  // We'll always reload current selected month.
  if (!activeMonthName) return;
  return reloadSingleMonth(activeMonthName);
}

async function reloadSingleMonth(tab) {
  if (!tab) return;
  resetViewState();

  const payload = await apiGet(
    `/api/getSheetData?sheet=${encodeURIComponent(tab)}`,
    `getSheetData:${tab}`
  ).catch((err) => {
    console.error("getSheetData failed:", err);
    return { headers: [], rows: [] };
  });

  currentHeaders = payload.headers || [];
  currentRows = Array.isArray(payload.rows) ? payload.rows : [];

  drawHeader(currentHeaders);
  renderTable();

  setSummaryLabel(tab, currentRows.length);
}

/* ========== Table header + sorting ========== */
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

/* ========== Table rendering (paginate) ========== */
function renderTable() {
  const rows = currentRows.slice();

  // Sort
  rows.sort((a, b) => {
    const va = a?.[sortCol];
    const vb = b?.[sortCol];

    const na = Number(va);
    const nb = Number(vb);

    const bothNum =
      !Number.isNaN(na) &&
      !Number.isNaN(nb) &&
      String(va).trim() !== "" &&
      String(vb).trim() !== "";

    const cmp = bothNum
      ? na - nb
      : String(va ?? "").toLowerCase() < String(vb ?? "").toLowerCase()
      ? -1
      : String(va ?? "").toLowerCase() > String(vb ?? "").toLowerCase()
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

      (row || []).forEach((cell) => {
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

/* ========== Timezone helpers ========== */
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
  const parts = new Intl.DateTimeFormat(undefined, { timeZoneName: "short" }).formatToParts(
    new Date()
  );
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
