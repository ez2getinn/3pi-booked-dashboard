// api/getSheetNames/index.js
const { graphGet, mustEnv } = require("../_shared/msGraph");

module.exports = async function (context, req) {
  try {
    const fileId = mustEnv("MS_EXCEL_FILE_ID");

    // Get worksheet list from the Excel workbook
    const data = await graphGet(
      `/me/drive/items/${encodeURIComponent(fileId)}/workbook/worksheets`
    );

    const names = (data.value || [])
      .map((w) => w?.name)
      .filter(Boolean);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: names,
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: err.message || String(err),
      },
    };
  }
};
