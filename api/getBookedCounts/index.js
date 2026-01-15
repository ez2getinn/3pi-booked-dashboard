// api/getBookedCounts/index.js
const { graphGet, mustEnv, resolveSiteAndDrive } = require("../_shared/msGraph");

/**
 * Month sheet detection:
 * Supports: Jan, January, Jan 2026, February 2026, etc.
 */
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

// ✅ Explicitly exclude non-month tabs
const EXCLUDED_TABS = new Set(["logs", "tech"]);

function parseMonthSheet(name) {
  const clean = String(name || "").trim();
  if (!clean) return null;

  const m = MONTH_RX.exec(clean);
  if (!m) return null;

  const key = m[1].slice(0, 3).toLowerCase();
  const monthIndex = MONTH_MAP[key];
  const year = m[2] ? parseInt(m[2], 10) : null;

  if (monthIndex == null) return null;
  return { monthIndex, year, raw: clean };
}

function normHeader(h) {
  return String(h ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function isBookedCell(v) {
  return String(v ?? "").trim().toUpperCase() === "BOOKED";
}

module.exports = async function (context, req) {
  try {
    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // 1) Get worksheet names
    const wsUrl =
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets?$select=name`;

    const ws = await graphGet(wsUrl);

    // 2) Filter only month sheets (exclude logs/tech)
    let sheetNames = (ws.value || [])
      .map((x) => x.name)
      .filter(Boolean)
      .map((n) => String(n).trim())
      .filter((n) => n.length > 0)
      .filter((n) => !EXCLUDED_TABS.has(n.toLowerCase()))
      .filter((n) => !!parseMonthSheet(n));

    // 3) Sort Jan->Dec, then by year
    sheetNames.sort((a, b) => {
      const pa = parseMonthSheet(a);
      const pb = parseMonthSheet(b);

      const ya = pa?.year ?? 9999;
      const yb = pb?.year ?? 9999;
      if (ya !== yb) return ya - yb;

      return (pa?.monthIndex ?? 0) - (pb?.monthIndex ?? 0);
    });

    const results = [];

    // 4) Count BOOKED rows for each month sheet
    for (const sheetName of sheetNames) {
      const safeSheet = sheetName.replace(/'/g, "''");

      const rangeUrl =
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
        `/workbook/worksheets('${safeSheet}')/usedRange(valuesOnly=true)?$select=values`;

      const range = await graphGet(rangeUrl);

      const values = Array.isArray(range.values) ? range.values : [];
      const headers = values[0] || [];
      const rows = values.slice(1);

      // ✅ find "BOOKED" column more safely
      const bookedIdx = headers.findIndex((h) => normHeader(h) === "booked");

      let count = 0;

      if (bookedIdx >= 0) {
        for (const r of rows) {
          const cell = r?.[bookedIdx];
          if (isBookedCell(cell)) count++;
        }
      }

      results.push({ name: sheetName, count });
    }

    // ✅ Force no-cache so new Feb updates appear immediately
    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store, no-cache, must-revalidate, proxy-revalidate",
        Pragma: "no-cache",
        Expires: "0",
      },
      body: results,
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: err?.message || String(err),
        name: err?.name,
        stack: err?.stack,
        response: err?.response?.data || err?.response || null,
      },
    };
  }
};
