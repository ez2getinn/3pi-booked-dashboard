const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  try {
    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) throw new Error("Missing query param: sheet");

    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // Escape single quotes for worksheet('...')
    const safeSheet = sheet.replace(/'/g, "''");

    const range = await graphGet(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
        `/workbook/worksheets('${encodeURIComponent(safeSheet)}')/usedRange(valuesOnly=true)?$select=values`
    );

    const values = range.values || [];
    if (!values.length) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: { sheet, headers: [], rows: [], ms: [] },
      };
      return;
    }

    const headers = values[0] || [];
    const rows = values.slice(1);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { sheet, headers, rows, ms: [] },
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: { ok: false, error: err.message || String(err) },
    };
  }
};
