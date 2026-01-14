async function getAccessToken() {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.append("client_id", clientId);
  body.append("client_secret", clientSecret);
  body.append("scope", "https://graph.microsoft.com/.default");
  body.append("grant_type", "client_credentials");

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  const json = await res.json();
  if (!res.ok) {
    throw new Error(`Token error ${res.status}: ${JSON.stringify(json)}`);
  }

  return json.access_token;
}

async function graphGet(url, accessToken) {
  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
    },
  });

  const text = await res.text();

  if (!res.ok) {
    throw new Error(`Graph GET failed ${res.status}: ${text}`);
  }

  return JSON.parse(text);
}

module.exports = async function (context, req) {
  try {
    const siteHost = process.env.MS_EXCEL_SITE_HOST;
    const sitePath = process.env.MS_EXCEL_SITE_PATH;
    const fileItemId = process.env.MS_EXCEL_FILE_ID;

    if (!siteHost || !sitePath || !fileItemId) {
      context.res = {
        status: 500,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: false,
          error:
            "Missing env vars. Need MS_EXCEL_SITE_HOST, MS_EXCEL_SITE_PATH, MS_EXCEL_FILE_ID",
        },
      };
      return;
    }

    const token = await getAccessToken();

    // ✅ Resolve SharePoint site
    const siteUrl = `https://graph.microsoft.com/v1.0/sites/${siteHost}:${sitePath}`;
    const site = await graphGet(siteUrl, token);

    // ✅ Get worksheet names
    const wsUrl = `https://graph.microsoft.com/v1.0/sites/${site.id}/drive/items/${fileItemId}/workbook/worksheets?$select=name`;
    const ws = await graphGet(wsUrl, token);

    const names = Array.isArray(ws?.value)
      ? ws.value.map((x) => x.name).filter(Boolean)
      : [];

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: names,
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
