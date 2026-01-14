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

// api/_shared/msGraph.js
const fetch = require("node-fetch");

function mustEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing env var: ${name}`);
  return v;
}

let tokenCache = { accessToken: null, exp: 0 };

async function getAppToken() {
  const tenant = mustEnv("MS_TENANT_ID");
  const clientId = mustEnv("MS_CLIENT_ID");
  const clientSecret = mustEnv("MS_CLIENT_SECRET");

  const now = Math.floor(Date.now() / 1000);
  if (tokenCache.accessToken && tokenCache.exp > now + 60) return tokenCache.accessToken;

  const url = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  });

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const json = await res.json();
  if (!res.ok) throw new Error(`Token error ${res.status}: ${JSON.stringify(json)}`);

  // expires_in is seconds
  tokenCache.accessToken = json.access_token;
  tokenCache.exp = now + (json.expires_in || 3599);
  return tokenCache.accessToken;
}

async function graphGet(path) {
  const token = await getAppToken();
  const url = `https://graph.microsoft.com/v1.0${path}`;

  const res = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}` },
  });

  const text = await res.text();
  let json;
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`Graph non-JSON ${res.status}: ${text.slice(0, 500)}`);
  }

  if (!res.ok) {
    throw new Error(`Graph GET failed ${res.status} ${path}: ${JSON.stringify(json)}`);
  }

  return json;
}

// Cache resolved IDs so we don't re-resolve every request
let resolved = { siteId: null, driveId: null };

async function resolveSiteAndDrive() {
  if (resolved.siteId && resolved.driveId) return resolved;

  const host = mustEnv("MS_EXCEL_SITE_HOST");   // itguru4u-my.sharepoint.com
  const sitePath = mustEnv("MS_EXCEL_SITE_PATH"); // /personal/... (must start with /)

  const site = await graphGet(`/sites/${host}:${sitePath}`);
  const siteId = site.id;
  if (!siteId) throw new Error("Could not resolve siteId from host+path");

  // Default document library drive for the site
  const drive = await graphGet(`/sites/${encodeURIComponent(siteId)}/drive`);
  const driveId = drive.id;
  if (!driveId) throw new Error("Could not resolve driveId from siteId");

  resolved = { siteId, driveId };
  return resolved;
}

module.exports = { mustEnv, graphGet, resolveSiteAndDrive };


