const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: { ok: true, step: "start" },
  };

  try {
    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) throw new Error("Missing query param: sheet");

    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // do NOT encode inside worksheets('...')
    const safeSheet = sheet.replace(/'/g, "''");

    const url =
      `/drives/${driveId}/items/${fileId}` +
      `/workbook/worksheets('${safeSheet}')/usedRange(valuesOnly=true)?$select=values`;

    // DEBUG: return URL first (so we know exactly what endpoint is hit)
    const range = await graphGet(url);

    const values = range?.values || [];

    if (!values.length) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: true,
          sheet,
          driveId,
          fileId,
          valuesCount: 0,
          headers: [],
          rows: [],
        },
      };
      return;
    }

    const headers = values[0] || [];
    const rows = values.slice(1);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        sheet,
        driveId,
        fileId,
        valuesCount: values.length,
        headers,
        rows,
      },
    };
  } catch (err) {
    // FORCE return full error details as JSON
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        name: err?.name,
        message: err?.message,
        stack: err?.stack,
        raw: String(err),
      },
    };
  }
};
