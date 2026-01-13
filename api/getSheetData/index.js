module.exports = async function (context, req) {
  try {
    const sheet = (req.query.sheet || "Sep").toString().trim();

    // ✅ Your GAS endpoint (the one that returns headers/rows/ms)
    // IMPORTANT: this endpoint MUST support ?sheet=Sep
    const GAS_BASE =
      "https://script.google.com/a/macros/shift4.com/s/AKfycbxsy7lAvTMPuZf9GW_ER4iuScDIb5vCUhJ0Rx4SA5Pu_1nXXHMSonErxfKj6bJ5T1mPZA/exec";

    const url = `${GAS_BASE}?sheet=${encodeURIComponent(sheet)}`;

    const res = await fetch(url, { method: "GET" });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          error: "GAS getSheetData failed",
          status: res.status,
          details: txt,
        },
      };
      return;
    }

    const data = await res.json();

    // ✅ Return exactly what frontend expects:
    // { sheet, headers, rows, ms }
    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store",
      },
      body: data,
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        error: "getSheetData exception",
        message: String(err?.message || err),
      },
    };
  }
};
