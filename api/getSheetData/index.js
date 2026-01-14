const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  try {
    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) throw new Error("Missing query param: sheet");

    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    // escape ' for worksheet('...')
    const safeSheet = sheet.replace(/'/g, "''");

    const url =
      `/drives/${driveId}/items/${fileId}` +
      `/workbook/worksheets('${safeSheet}')/usedRange(valuesOnly=true)?$select=values`;

    // LOG URL
    context.log("DEBUG Graph URL:", url);

    const range = await graphGet(url);

    const values = range.values || [];

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        sheet,
        debug: { driveId, fileId, url },
        valuesPreview: values.slice(0, 3),
        valuesCount: values.length,
      },
    };
  } catch (err) {
    // IMPORTANT: show EVERYTHING
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        message: err.message || String(err),
        stack: err.stack || null,
        raw: err,
      },
    };
  }
};
