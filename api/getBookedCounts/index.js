const { graphGet, mustEnv, resolveSiteAndDrive } = require("../_shared/msGraph");

/* ✅ Month detection (Jan / January / Jan 2026) */
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

function parseMonthSheet(name) {
  const m = MONTH_RX.exec(String(name || "").trim());
  if (!m) return null;

  const key = m[1].slice(0, 3).toLowerCase();
  const monthIndex = MONTH_MAP[key];
  const year = m[2] ? parseInt(m[2], 10) : null;

  if (monthIndex == null) return null;
  return { monthIndex, year };
}

function normHeader(h) {
  return String(h || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

module.exports = async function (context, req) {
  try {
    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // 1) worksheet list
    const wsUrl =
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets?$select=name`;

    const ws = await graphGet(wsUrl);

    // ✅ keep only month sheets
    let sheetNames = (ws.value || [])
      .map((x) => x.name)
      .filter(Boolean)
      .filter((n) => !!parseMonthSheet(n));

    // ✅ sort Jan → Dec
    sheetNames.sort((a, b) => {
      const pa = parseMonthSheet(a);
      const pb = parseMonthSheet(b);

      const ya = pa.year ?? 9999;
      const yb = pb.year ?? 9999;
      if (ya !== yb) return ya - yb;

      return pa.monthIndex - pb.monthIndex;
    });

    const results = [];

    for (const sheetName of sheetNames) {
      const safeSheet = sheetName.replace(/'/g, "''");

      const rangeUrl =
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
        `/workbook/worksheets('${safeSheet}')/usedRange(valuesOnly=true)?$select=values`;

      const range = await graphGet(rangeUrl);
      const values = range.values || [];

      // ✅ If sheet empty
      if (!values.length) {
        results.push({ name: sheetName, count: 0 });
        continue;
      }

      // ✅ Find first non-empty header row (fixes Excel usedRange header misalignment)
      let headerRowIndex = -1;
      for (let i = 0; i < Math.min(values.length, 5); i++) {
        const row = values[i] || [];
        const hasAnyHeaderText = row.some((c) => normHeader(c) !== "");
        if (hasAnyHeaderText) {
          headerRowIndex = i;
          break;
        }
      }

      if (headerRowIndex < 0) {
        results.push({ name: sheetName, count: 0 });
        continue;
      }

      const headers = values[headerRowIndex] || [];
      const rows = values.slice(headerRowIndex + 1);

      // ✅ Find booked column more flexibly
      const bookedIdx = headers.findIndex((h) => normHeader(h) === "booked");

      let count = 0;

      if (bookedIdx >= 0) {
        for (const r of rows) {
          const cell = r?.[bookedIdx];
          if (String(cell || "").trim().toUpperCase() === "BOOKED") count++;
        }
      }

      results.push({ name: sheetName, count });
    }

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store",
      },
      body: results,
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store",
      },
      body: {
        ok: false,
        error: err?.message || String(err),
        name: err?.name,
        stack: err?.stack,
      },
    };
  }
};
