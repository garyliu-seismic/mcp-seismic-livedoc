import * as fs from "fs";
import * as path from "path";

// Reads coworkUserFilesPath from the Claude Desktop config JSON.
// Tries the standard %APPDATA%\Claude path first, then the Microsoft Store
// package path (%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude).
export function findCoworkPath(): string {
  const candidates: string[] = [];
  const appdata = process.env.APPDATA;
  const localAppdata = process.env.LOCALAPPDATA;
  if (appdata) candidates.push(path.join(appdata, "Claude", "claude_desktop_config.json"));
  if (localAppdata) {
    const pkgsDir = path.join(localAppdata, "Packages");
    try {
      for (const entry of fs.readdirSync(pkgsDir)) {
        if (entry.startsWith("Claude_")) {
          candidates.push(path.join(pkgsDir, entry, "LocalCache", "Roaming", "Claude", "claude_desktop_config.json"));
        }
      }
    } catch { /* Packages dir unreadable */ }
  }
  for (const cfgPath of candidates) {
    try {
      const cfg = JSON.parse(fs.readFileSync(cfgPath, "utf8")) as Record<string, unknown>;
      const p = cfg.coworkUserFilesPath;
      if (typeof p === "string" && p) return p;
    } catch { /* not found or not parseable */ }
  }
  return "";
}

export const COWORK_PATH = process.env.CLAUDE_COWORK_PATH || findCoworkPath();

export const BASE_URL = process.env.SEISMIC_BASE_URL ?? "https://api.seismic.com/livedoc";

// The Integration API is a sibling service of the LiveDoc API, used to resolve a DocCenter
// profile's id/versionId from its name (Document Generator's Public API has no such lookup).
// Its gateway route is NOT a simple sibling path of LiveDoc's — per
// api-specifications/packages/apim-route/public-api-mapping-{prod,non-prod}.yml:
//   prod:            LiveDoc "/livedoc"        Integration "/integration"
//   dev/qa/uat:       LiveDoc "/{env}/livedoc"  Integration "/{env}/services/integration"
// (non-prod inserts an extra "services/" segment that prod does not have), so this must be
// derived per-environment rather than by string-editing BASE_URL's trailing segment.
function deriveIntegrationBaseUrl(baseUrl: string): string {
  const url = new URL(baseUrl);
  const nonProdMatch = url.pathname.match(/^\/(dev|qa|uat)\/livedoc\/?$/);
  if (nonProdMatch) {
    return `${url.origin}/${nonProdMatch[1]}/services/integration`;
  }
  return `${url.origin}/integration`;
}

export const INTEGRATION_BASE_URL = process.env.SEISMIC_INTEGRATION_BASE_URL ?? deriveIntegrationBaseUrl(BASE_URL);

// The workspace-browsing endpoints (GetWorkspaceDestinationSpaces/Roots/RootItems/FolderItems) live
// in a separate "Document Generator (Internal)" API resource — added 2026-09-03 with its own api_id
// so it can't collide with the main LiveDoc api_id (see api-specifications commit 06dc51c1). Unlike
// the Integration API, this one IS a simple sibling path: LiveDoc "/{env}/livedoc" (prod: "/livedoc")
// -> Internal "/{env}/livedoc-internal" (prod: "/livedoc-internal"), no extra segment inserted.
function deriveInternalBaseUrl(baseUrl: string): string {
  const url = new URL(baseUrl);
  return `${url.origin}${url.pathname.replace(/\/livedoc\/?$/, "/livedoc-internal")}`;
}

export const INTERNAL_BASE_URL = process.env.SEISMIC_INTERNAL_BASE_URL ?? deriveInternalBaseUrl(BASE_URL);

export const DEFAULT_AUTH_URI    = process.env.AUTH_SERVICE_URI ?? "https://auth-qa.seismic-dev.com";
export const DEFAULT_AUTH_TENANT = process.env.AUTH_TENANT      ?? "";
export const DEFAULT_USERNAME    = process.env.AUTH_USERNAME    ?? "";
export const DEFAULT_PASSWORD    = process.env.AUTH_PASSWORD    ?? "";

// Public DocCenter web client — works across all tenants, no secret required.
export const BROWSER_CLIENT_ID = "0188a34d-cdbd-4208-8ebc-d0567984915e";
export const BROWSER_SCOPES    = "openid id library download engagement_read engagement_write upload " +
  "feature_read collection_read contentdiscovery doccenter_backend_read livedoc " +
  "ums_bff_read das_data_rw email profile entitlement_read aiml_llm";

export const FORM_API_BASE = process.env.FORM_API_URL ?? "http://localhost:3001";
export const FORM_RESOURCE_URI = "ui://livedoc/form";
