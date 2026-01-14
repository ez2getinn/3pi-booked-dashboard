// api/getSheetData/index.js
const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  try {
    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) throw new Error("Missing query param: sheet");

    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // ✅ Graph expects raw worksheet name inside ('...')
    // We ONLY escape single quotes (Excel worksheet name rules)
    const safeSheetName = sheet.replace(/'/g, "''");

    // ✅ DO NOT encodeURIComponent(sheet) inside the worksheet('...')
    // Only encode parts like driveId/fileId (path segments)
    const url =
      `/drives/${encodeURIComponent(driveId)}` +
      `/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets('${safeSheetName}')/usedRange(valuesOnly=true)?$select=values`;

    const range = await graphGet(url);

    const values = range?.values || [];

    if (!values.length) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: { sheet, headers: [], rows: [] }
      };
      return;
    }

    const headers = values[0] || [];
    const rows = values.slice(1);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { sheet, headers, rows }
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: err?.message || String(err),
        // ✅ extra debug info to see real cause in browser
        stack: err?.stack || null
      }
    };
  }
};
