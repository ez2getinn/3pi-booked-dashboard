const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  try {
    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) throw new Error("Missing query param: sheet");

    // ✅ ENV values
    const tenantId = mustEnv("MS_TENANT_ID");
    const clientId = mustEnv("MS_CLIENT_ID");
    const fileId = mustEnv("MS_EXCEL_FILE_ID");

    const resolved = await resolveSiteAndDrive(); // { siteId, driveId }
    const driveId = resolved.driveId;

    // ✅ Sheet name MUST be raw inside ('...')
    const safeSheetName = sheet.replace(/'/g, "''");

    const url =
      `/drives/${encodeURIComponent(driveId)}` +
      `/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets('${safeSheetName}')/usedRange(valuesOnly=true)?$select=values`;

    // ✅ call Graph
    const range = await graphGet(url);
    const values = range?.values || [];

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        sheet,
        debug: {
          tenantId,
          clientId,
          fileId,
          driveId,
          graphUrl: url,
        },
        valuesCount: values.length,
        headersRow: values[0] || [],
        rowsCount: Math.max(values.length - 1, 0),
      },
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: err?.message || String(err),
        stack: err?.stack || null,
      },
    };
  }
};
