// api/getSheetNames/index.js
// Microsoft Graph version - returns Excel worksheet names

const fetch = require("node-fetch");

async function getAccessToken() {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET in env vars");
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.append("client_id", clientId);
  body.append("client_secret", clientSecret);
  body.append("grant_type", "client_credentials");
  body.append("scope", "https://graph.microsoft.com/.default");

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  const json = await res.json();

  if (!res.ok) {
    throw new Error(`Token error: ${JSON.stringify(json)}`);
  }

  return json.access_token;
}

module.exports = async function (context, req) {
  try {
    const siteHost = process.env.MS_EXCEL_SITE_HOST; // ex: itguru4u-my.sharepoint.com
    const fileId = process.env.MS_EXCEL_FILE_ID;     // ex: 9C200944-0FEC-444B-AAF9-FDC864ED5B54

    if (!siteHost || !fileId) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: false,
          error: "Missing MS_EXCEL_SITE_HOST or MS_EXCEL_FILE_ID in env vars",
        },
      };
      return;
    }

    const token = await getAccessToken();

    // ✅ We need Site ID first
    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteHost}:/personal/ashiryousaf_itguru4u_onmicrosoft_com`,
      {
        method: "GET",
        headers: { Authorization: `Bearer ${token}` },
      }
    );

    const siteJson = await siteRes.json();
    if (!siteRes.ok) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: false,
          error: "Failed to get Site ID",
          details: siteJson,
        },
      };
      return;
    }

    const siteId = siteJson.id;

    // ✅ List worksheets
    const wsRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/workbook/worksheets`,
      {
        method: "GET",
        headers: { Authorization: `Bearer ${token}` },
      }
    );

    const wsJson = await wsRes.json();
    if (!wsRes.ok) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: false,
          error: "Failed to get worksheet names",
          details: wsJson,
        },
      };
      return;
    }

    const sheetNames = (wsJson.value || []).map((w) => w.name);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: sheetNames,
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: "getSheetNames exception",
        message: String(err.message || err),
      },
    };
  }
};
