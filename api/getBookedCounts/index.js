module.exports = async function (context, req) {
  try {
    // ✅ Your GAS endpoint already returns JSON like:
    // [ { "name": "Jan", "count": 26 }, { "name": "Feb", "count": 0 }, ... ]
    const GAS_URL =
      "https://script.google.com/a/macros/shift4.com/s/AKfycbxb4baiuXwUx0UGj9r79eFVhijOLN0fX3dHSbClYLeVM_AhZSW00uzntZDWGi0iMLIqyA/exec";

    const res = await fetch(GAS_URL, { method: "GET" });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: { error: "GAS getBookedCounts failed", status: res.status, details: txt },
      };
      return;
    }

    const data = await res.json();

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
        error: "getBookedCounts exception",
        message: String(err?.message || err),
      },
    };
  }
};
