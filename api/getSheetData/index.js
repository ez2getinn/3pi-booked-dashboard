module.exports = async function (context, req) {
  try {
    const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) {
      context.res = { status: 400, body: { ok: false, error: "Missing query param: sheet" } };
      return;
    }

    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const tenantId = mustEnv("MS_TENANT_ID");
    const siteHost = mustEnv("MS_EXCEL_SITE_HOST");
    const sitePath = mustEnv("MS_EXCEL_SITE_PATH");

    const { siteId, driveId } = await resolveSiteAndDrive();

    const safeSheet = sheet.replace(/'/g, "''");

    const url =
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets('${safeSheet}')/usedRange(valuesOnly=true)?$select=values`;

    const range = await graphGet(url);

    const values = range.values || [];
    const headers = values[0] || [];
    const rows = values.slice(1);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        sheet,
        headers,
        rows,
        debug: {
          tenantId,
          siteHost,
          sitePath,
          siteId,
          driveId,
          fileId,
          graphUrl: url,
        },
      },
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
