const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  context.log("🔥 getSheetData called");
  context.log("Query:", req.query);

  try {
    const sheet = String(req.query.sheet || "").trim();
    if (!sheet) throw new Error("Missing query param: sheet");

    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    context.log("MS_EXCEL_FILE_ID:", fileId);

    const { driveId, siteId } = await resolveSiteAndDrive();
    context.log("Resolved siteId:", siteId);
    context.log("Resolved driveId:", driveId);

    const safeSheet = sheet.replace(/'/g, "''");
    const SAFE_RANGE = "A1:Z50";

    const url =
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}` +
      `/workbook/worksheets('${safeSheet}')/range(address='${SAFE_RANGE}')?$select=values`;

    context.log("Graph URL:", url);

    const range = await graphGet(url);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        sheet,
        range: SAFE_RANGE,
        valuesPreview: (range.values || []).slice(0, 5),
      },
    };
  } catch (err) {
    // ✅ RETURN FULL ERROR DETAILS
    context.log.error("❌ ERROR:", err);

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
