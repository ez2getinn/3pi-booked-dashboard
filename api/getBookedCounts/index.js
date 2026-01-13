// api/getBookedCounts/index.js
// Reads BOOKED counts from your Google Apps Script Web App endpoint
// Returns: [{ name: "Jan", count: 26 }, ...] (JSON)

// ✅ PUT YOUR GAS WEB APP URL HERE (must return JSON array)
const GAS_GET_BOOKED_COUNTS_URL =
  "https://script.google.com/a/macros/shift4.com/s/AKfycbxb4baiuXwUx0UGj9r79eFVhijOLN0fX3dHSbClYLeVM_AhZSW00uzntZDWGi0iMLIqyA/exec";

module.exports = async function (context, req) {
  try {
    // Call GAS endpoint
    const res = await fetch(GAS_GET_BOOKED_COUNTS_URL, {
      method: "GET",
      headers: {
        "Accept": "application/json",
      },
    });

    const text = await res.text();

    // If GAS returns an HTML login page, this will catch it
    if (!res.ok) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          error: "getBookedCounts exception",
          message: `GAS returned HTTP ${res.status}`,
          preview: text.slice(0, 200),
        },
      };
      return;
    }

    // Parse JSON
    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          error: "getBookedCounts exception",
          message: `Unexpected token in response. Not valid JSON.`,
          preview: text.slice(0, 200),
        },
      };
      return;
    }

    // Must be an array like: [{name,count}, ...]
    if (!Array.isArray(data)) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          error: "getBookedCounts exception",
          message: "GAS response is not an array",
          receivedType: typeof data,
          received: data,
        },
      };
      return;
    }

    // Clean/normalize
    const cleaned = data.map((x) => ({
      name: String(x?.name ?? "").trim(),
      count: Number(x?.count ?? 0) || 0,
    }));

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: cleaned,
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
