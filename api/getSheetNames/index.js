module.exports = async function (context, req) {
  try {
    const base = process.env.GAS_API_BASE;
    if (!base) {
      context.res = { status: 500, body: "Missing GAS_API_BASE secret" };
      return;
    }

    // calls your Apps Script function getSheetNames
    const url = `${base}?fn=getSheetNames`;
    const r = await fetch(url);
    const data = await r.json();

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: data,
    };
  } catch (e) {
    context.res = { status: 500, body: String(e) };
  }
};
