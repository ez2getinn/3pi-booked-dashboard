const { mustEnv, graphGet, resolveSiteAndDrive } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  try {
    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const { driveId } = await resolveSiteAndDrive();

    const data = await graphGet(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(fileId)}/workbook/worksheets`
    );

    const names = (data.value || []).map((w) => w.name).filter(Boolean);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: names,
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: { ok: false, error: err.message || String(err) },
    };
  }
};
