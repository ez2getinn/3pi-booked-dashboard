module.exports = async function (context, req) {
  try {
    // ✅ GAS endpoint (your working web app)
    const GAS_BASE =
      "https://script.google.com/a/macros/shift4.com/s/AKfycbxsy7lAvTMPuZf9GW_ER4iuScDIb5vCUhJ0Rx4SA5Pu_1nXXHMSonErxfKj6bJ5T1mPZA/exec";

    // ✅ This will return JSON like:
    // [{ "name":"Jan","count":26 }, ...]
    const url = `${GAS_BASE}?fn=getBookedCounts`;

    const res = await fetch(url, { method: "GET" });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          error: "GAS getBookedCounts failed",
          status: res.status,
          details: txt,
        },
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
      body: { error: "getBookedCounts exception", message: String(err?.message || err) },
    };
  }
};
