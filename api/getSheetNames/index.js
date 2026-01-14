module.exports = async function (context, req) {
  try {
    // ✅ TEMP: Hardcoded tab list (so frontend works instantly)
    // Later we will read this list dynamically from Excel
    const tabs = ["Sep 2025", "Oct 2025", "Nov 2025", "Dec 2025"];

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: tabs
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: false,
        error: err.message || String(err)
      }
    };
  }
};
