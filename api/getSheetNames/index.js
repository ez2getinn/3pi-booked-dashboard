// api/getSheetNames/index.js
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

module.exports = async function (context, req) {
  try {
    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // ✅ List worksheets from the correct drive+file
    const wsUrl =
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets?$select=name`;

    const ws = await graphGet(wsUrl);

    // ✅ keep only real month sheets (exclude logs/tech/etc)
    let sheetNames = (ws.value || [])
      .map((x) => x.name)
      .filter(Boolean)
      .filter((n) => !!parseMonthSheet(n));

    // ✅ sort by year (if present), then month Jan→Dec
    sheetNames.sort((a, b) => {
      const pa = parseMonthSheet(a);
      const pb = parseMonthSheet(b);

      const ya = pa.year ?? 9999;
      const yb = pb.year ?? 9999;
      if (ya !== yb) return ya - yb;

      return pa.monthIndex - pb.monthIndex;
    });

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store",
      },
      body: sheetNames,
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
