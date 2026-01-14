// api/getSheetData/index.js
const { graphGet, mustEnv } = require("../_shared/msGraph");

function normalizeHeader(h) {
  return String(h || "").trim();
}

function parseDateToMs(v) {
  // Excel may return:
  // - "1/10/2026"
  // - "2026-01-10"
  // - a number (Excel serial date)
  if (v == null || v === "") return null;

  // If it's already a number, treat it like Excel date serial
  // Excel serial date: days since 1899-12-30
  if (typeof v === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return excelEpoch.getTime() + v * 86400000;
  }

  const s = String(v).trim();
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getTime();

  return null;
}

function parseTimeToMs(dateMs, timeVal) {
  if (!dateMs || timeVal == null || timeVal === "") return null;

  // If time is number = fraction of a day in Excel (ex: 0.5 = 12:00)
  if (typeof timeVal === "number") {
    return dateMs + Math.round(timeVal * 86400000);
  }

  const s = String(timeVal).trim();

  // Try parse "8:00 AM"
  const d = new Date(dateMs);
  const t = new Date(`${d.toDateString()} ${s}`);
  if (!isNaN(t.getTime())) return t.getTime();

  return null;
}

module.exports = async function (context, req) {
  try {
    const fileId = mustEnv("MS_EXCEL_FILE_ID");
    const sheetName = (req.query.sheet || "").trim();

    if (!sheetName) {
      throw new Error("Missing query param: sheet");
    }

    // Read used range of worksheet
    // This returns { values: [ [headers...], [row...], ... ] }
    const range = await graphGet(
      `/me/drive/items/${encodeURIComponent(
        fileId
      )}/workbook/worksheets('${encodeURIComponent(
        sheetName.replace(/'/g, "''")
      )}')/usedRange(valuesOnly=true)?$select=values`
    );

    const values = range.values || [];
    if (!values.length) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: { sheet: sheetName, headers: [], rows: [], ms: [] },
      };
      return;
    }

    const headers = values[0].map(normalizeHeader);
    const rows = values.slice(1).filter((r) => Array.isArray(r) && r.some((x) => String(x || "").trim() !== ""));

    // find key columns by name (your frontend expects these indices)
    // fallback to fixed indexes if names match your layout
    const colDate = headers.findIndex((h) => h.toLowerCase() === "date");
    const colStart = headers.findIndex((h) => h.toLowerCase().includes("start"));
    const colEnd = headers.findIndex((h) => h.toLowerCase().includes("end"));

    const ms = rows.map((r) => {
      const dateVal = colDate >= 0 ? r[colDate] : null;
      const dateMs = parseDateToMs(dateVal);

      const startVal = colStart >= 0 ? r[colStart] : null;
      const endVal = colEnd >= 0 ? r[colEnd] : null;

      return {
        dateMs,
        startMs: parseTimeToMs(dateMs, startVal),
        endMs: parseTimeToMs(dateMs, endVal),
      };
    });

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        sheet: sheetName,
        headers,
        rows,
        ms,
      },
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
