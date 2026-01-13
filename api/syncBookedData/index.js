module.exports = async function (context, req) {
  try {
    // ✅ Security check (optional but recommended)
    const providedKey = (req.query.key || "").toString();
    const expectedKey = (process.env.SYNC_KEY || "").toString();

    if (!expectedKey) {
      context.res = {
        status: 500,
        body: { ok: false, error: "SYNC_KEY missing in Azure env variables" },
      };
      return;
    }

    if (providedKey !== expectedKey) {
      context.res = {
        status: 401,
        body: { ok: false, error: "Unauthorized. Missing/invalid key." },
      };
      return;
    }

    // ✅ GAS URL stored in Azure env variables
    const gasUrl = (process.env.GAS_SYNC_URL || "").toString().trim();

    if (!gasUrl) {
      context.res = {
        status: 500,
        body: { ok: false, error: "GAS_SYNC_URL missing in Azure env variables" },
      };
      return;
    }

    // ✅ Call your Google Apps Script web app
    const res = await fetch(gasUrl, { method: "GET" });

    const text = await res.text();

    // ✅ GAS should return JSON
    let json;
    try {
      json = JSON.parse(text);
    } catch (err) {
      context.res = {
        status: 500,
        body: {
          ok: false,
          error: "GAS response is not valid JSON",
          preview: text.substring(0, 300),
        },
      };
      return;
    }

    // ✅ Return the response back to browser (Azure API)
    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        source: "GAS",
        gasResponse: json,
      },
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { ok: false, error: "syncBookedData exception", message: String(err) },
    };
  }
};
