// api/_shared/msGraph.js
// Shared helpers for Microsoft Graph (client_credentials)

async function getAccessToken() {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET");
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("grant_type", "client_credentials");
  body.set("scope", "https://graph.microsoft.com/.default");

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const json = await res.json();
  if (!res.ok) {
    throw new Error(
      `Token error: ${res.status} ${JSON.stringify(json).slice(0, 500)}`
    );
  }

  return json.access_token;
}

async function graphGet(path) {
  const token = await getAccessToken();

  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  const text = await res.text();

  let json;
  try {
    json = JSON.parse(text);
  } catch {
    json = { raw: text };
  }

  if (!res.ok) {
    throw new Error(
      `Graph GET failed: ${res.status} ${path} -> ${JSON.stringify(json).slice(
        0,
        500
      )}`
    );
  }

  return json;
}

function mustEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing environment variable: ${name}`);
  return v;
}

module.exports = {
  graphGet,
  mustEnv,
};
