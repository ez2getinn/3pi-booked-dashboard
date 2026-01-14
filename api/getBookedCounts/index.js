const { graphGet, mustEnv, resolveSiteAndDrive } = require("../_shared/msGraph");

/* ✅ Only allow month worksheet names:
   - Jan / January
   - Feb / February
   - ...
   - Dec / December
   - Optional: "Jan 2026"
*/
const MONTH_RX =
  /^(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)(?:\s+\d{4})?$/i;

function isMonthSheet(name) {
  return MONTH_RX.test(String(name || "").trim());
}

module.exports = async function (context, req) {
  try {
    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // 1) Get worksheet list
    const wsUrl =
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets?$select=name`;

    const ws = await graphGet(wsUrl);

    // ✅ filter only month sheets
    const names = (ws.value || [])
      .map((x) => x.name)
      .filter(Boolean)
      .filter(isMonthSheet);

    const results = [];

    for (const sheetName of names) {
      const safeSheet = sheetName.replace(/'/g, "''");

      const rangeUrl =
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
        `/workbook/worksheets('${safeSheet}')/usedRange(valuesOnly=true)?$select=values`;

      const range = await graphGet(rangeUrl);

      const values = range.values || [];
      const headers = values[0] || [];
      const rows = values.slice(1);

      // find BOOKED column index
      const bookedIdx = headers.findIndex(
        (h) => String(h || "").trim().toLowerCase() === "booked"
      );

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
      headers: { "Content-Type": "application/json" },
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
