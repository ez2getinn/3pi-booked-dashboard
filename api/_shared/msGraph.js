// api/_shared/msGraph.js
// Shared helpers for Microsoft Graph (client_credentials)

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
  if (tokenCache.accessToken && tokenCache.exp > now + 60) {
    return tokenCache.accessToken;
  }

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

  tokenCache.accessToken = json.access_token;
  tokenCache.exp = now + (json.expires_in || 3599);
  return tokenCache.accessToken;
}

async function graphGet(path) {
  const token = await getAppToken();
  const url = `https://graph.microsoft.com/v1.0${path}`;

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  const text = await res.text();
  let json = {};
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`Graph non-JSON ${res.status}: ${text.slice(0, 500)}`);
  }

  if (!res.ok) {
    throw new Error(`Graph GET failed ${res.status} ${path}: ${JSON.stringify(json).slice(0, 500)}`);
  }

  return json;
}

let resolved = { siteId: null, driveId: null };

async function resolveSiteAndDrive() {
  if (resolved.siteId && resolved.driveId) return resolved;

  const host = mustEnv("MS_EXCEL_SITE_HOST");
  const sitePath = mustEnv("MS_EXCEL_SITE_PATH");

  const site = await graphGet(`/sites/${host}:${sitePath}`);
  const siteId = site.id;
  if (!siteId) throw new Error("Could not resolve siteId from host+path");

  const drive = await graphGet(`/sites/${encodeURIComponent(siteId)}/drive`);
  const driveId = drive.id;
  if (!driveId) throw new Error("Could not resolve driveId from siteId");

  resolved = { siteId, driveId };
  return resolved;
}

module.exports = { mustEnv, graphGet, resolveSiteAndDrive };
