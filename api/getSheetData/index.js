const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  try {
    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) throw new Error("Missing query param: sheet");

    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // Graph needs worksheet('NAME') EXACTLY.
    // Only escape single quotes by doubling them.
    const safeSheet = sheet.replace(/'/g, "''");

    const url =
      `/drives/${driveId}/items/${fileId}` +
      `/workbook/worksheets('${safeSheet}')/usedRange(valuesOnly=true)?$select=values`;

    const range = await graphGet(url);

    const values = range.values || [];

    if (!values.length) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: { sheet, headers: [], rows: [] },
      };
      return;
    }

    const headers = values[0] || [];
    const rows = values.slice(1);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { sheet, headers, rows },
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: { ok: false, error: err.message || String(err) },
    };
  }
};
