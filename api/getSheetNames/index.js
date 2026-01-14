// api/getSheetNames/index.js
module.exports = async function (context, req) {
  try {
    const tenantId = process.env.MS_TENANT_ID;
    const clientId = process.env.MS_CLIENT_ID;
    const clientSecret = process.env.MS_CLIENT_SECRET;

    const siteHost = process.env.MS_EXCEL_SITE_HOST; // ex: wasa001-my.sharepoint.com
    const sitePath = process.env.MS_EXCEL_SITE_PATH; // ex: /personal/shahnazashir_wasa001_onmicrosoft_com
    const fileId = process.env.MS_EXCEL_FILE_ID;     // ex: A8CF92E4-3946-46D0-8F65-988DB0939B46

    if (!tenantId || !clientId || !clientSecret) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: { ok: false, error: "Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET" }
      };
      return;
    }

    if (!siteHost || !sitePath || !fileId) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: { ok: false, error: "Missing MS_EXCEL_SITE_HOST / MS_EXCEL_SITE_PATH / MS_EXCEL_FILE_ID" }
      };
      return;
    }

    // 1) Get app-only token
    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default"
      })
    });

    const tokenJson = await tokenRes.json();
    if (!tokenRes.ok) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: { ok: false, error: "Token request failed", details: tokenJson }
      };
      return;
    }

    const accessToken = tokenJson.access_token;

    // 2) Resolve SharePoint Site ID
    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteHost}:${sitePath}`,
      {
        headers: { Authorization: `Bearer ${accessToken}` }
      }
    );

    const siteJson = await siteRes.json();
    if (!siteRes.ok) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: { ok: false, error: "Failed to resolve SharePoint site", details: siteJson }
      };
      return;
    }

    const siteId = siteJson.id;

    // 3) Get Worksheets
    const wsRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/workbook/worksheets?$select=name`,
      {
        headers: { Authorization: `Bearer ${accessToken}` }
      }
    );

    const wsJson = await wsRes.json();
    if (!wsRes.ok) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: { ok: false, error: "Failed to fetch worksheet list", details: wsJson }
      };
      return;
    }

    const names = (wsJson.value || [])
      .map(w => w?.name)
      .filter(Boolean);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: names
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: { ok: false, error: err?.message || String(err) }
    };
  }
};
