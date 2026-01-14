const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  try {
    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) throw new Error("Missing query param: sheet");

    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // Escape worksheet name for Graph
    const safeSheet = sheet.replace(/'/g, "''");

    // ✅ IMPORTANT:
    // DO NOT use usedRange() because logs sheet / formatting makes it massive and causes 500.
    // Pull a SAFE range only.
    // Change range size if you want bigger later.
    const SAFE_RANGE = "A1:Z500";

    const url =
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets('${safeSheet}')/range(address='${SAFE_RANGE}')?$select=values`;

    const range = await graphGet(url);

    const values = range.values || [];

    // If empty
    if (!values.length) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: { sheet, range: SAFE_RANGE, headers: [], rows: [] },
      };
      return;
    }

    const headers = values[0] || [];
    const rows = values.slice(1).filter(r => r.some(cell => cell !== null && cell !== ""));

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { sheet, range: SAFE_RANGE, headers, rows },
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: { ok: false, error: err.message || String(err) },
    };
  }
};
