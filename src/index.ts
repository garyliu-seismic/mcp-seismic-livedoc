#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerAppTool, registerAppResource, RESOURCE_MIME_TYPE, getUiCapability } from "@modelcontextprotocol/ext-apps/server";
import { build } from "esbuild";
import { z } from "zod";
import * as http from "http";
import { randomUUID } from "crypto";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { exec } from "child_process";
import { fileURLToPath } from "url";

// Bundle @modelcontextprotocol/ext-apps (App class + deps) into a browser IIFE at startup.
// This avoids relying on Claude Desktop to inject an import map for bare specifiers.
let _extAppsBundleCache: string | null = null;
async function getExtAppsBundle(): Promise<string> {
  if (_extAppsBundleCache !== null) return _extAppsBundleCache;
  try {
    // resolveDir must be the package root (parent of node_modules).
    // dist/index.js → dist/ → package root; src/index.ts → src/ → package root.
    const pkgRoot = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
    const result = await build({
      stdin: {
        contents: `import { App } from "@modelcontextprotocol/ext-apps"; globalThis.__McpApp = { App };`,
        resolveDir: pkgRoot,
      },
      bundle: true,
      format: "iife",
      write: false,
      platform: "browser",
      logLevel: "silent",
    });
    _extAppsBundleCache = Buffer.from(result.outputFiles[0].contents).toString("utf-8");
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"),
      `[${new Date().toISOString()}] ext-apps bundle: ${_extAppsBundleCache.length} bytes OK\n`);
  } catch (e) {
    _extAppsBundleCache = `console.error("[livedoc] ext-apps bundle failed:", ${JSON.stringify(String(e))});`;
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"),
      `[${new Date().toISOString()}] ext-apps bundle FAILED: ${e}\n`);
  }
  return _extAppsBundleCache;
}

// Reads coworkUserFilesPath from the Claude Desktop config JSON.
// Tries the standard %APPDATA%\Claude path first, then the Microsoft Store
// package path (%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude).
function findCoworkPath(): string {
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

const COWORK_PATH = process.env.CLAUDE_COWORK_PATH || findCoworkPath();

const BASE_URL = process.env.SEISMIC_BASE_URL ?? "https://api.seismic.com/livedoc";

const DEFAULT_AUTH_URI      = process.env.AUTH_SERVICE_URI    ?? "";
const DEFAULT_AUTH_TENANT   = process.env.AUTH_TENANT         ?? "";
const DEFAULT_CLIENT_ID     = process.env.AUTH_CLIENT_ID      ?? "";
const DEFAULT_CLIENT_SECRET = process.env.AUTH_CLIENT_SECRET  ?? "";
const DEFAULT_USERNAME      = process.env.AUTH_USERNAME        ?? "";
const DEFAULT_PASSWORD      = process.env.AUTH_PASSWORD        ?? "";

// ── Tiny HTTP server for artifact form submissions ──────────────────────────
const FORM_PORT = 3099;
const pendingForms = new Map<string, (payload: unknown) => void>();
const pendingFormHtml = new Map<string, string>(); // token → full HTML, served via GET /form/<token>
// Submissions that arrived before wait_for_form_submit was called (race-condition buffer).
const preReceivedPayloads = new Map<string, unknown>();

const formHttpServer = http.createServer((req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }

  // Debug ping: GET /ping/<message> — logs from the App panel iframe for diagnostics.
  const pingMatch = req.url?.match(/^\/ping\/(.+)/);
  if (req.method === "GET" && pingMatch) {
    const msg = decodeURIComponent(pingMatch[1]);
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"), `[${new Date().toISOString()}] APP_PANEL: ${msg}\n`);
    res.writeHead(200, { "Content-Type": "text/plain" }); res.end("ok");
    return;
  }

  // Serve form HTML so a small <iframe> wrapper can load it inside Claude Cowork.
  const getForm = req.url?.match(/^\/form\/([^/?]+)/);
  if (req.method === "GET" && getForm) {
    const html = pendingFormHtml.get(getForm[1]);
    fs.appendFileSync(require("path").join(require("os").tmpdir(), "mcp-livedoc-debug.log"), `[${new Date().toISOString()}] HTTP GET /form/${getForm[1]}: found=${html !== undefined}\n`);
    if (html) {
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      res.end(html);
    } else {
      res.writeHead(404); res.end("Form not found or expired");
    }
    return;
  }

  const m = req.url?.match(/^\/submit\/([^/?]+)/);
  if (req.method === "POST" && m) {
    const token = m[1];
    let body = "";
    req.on("data", (chunk: Buffer) => { body += chunk; });
    req.on("end", () => {
      const resolver = pendingForms.get(token);
      fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"), `[${new Date().toISOString()}] HTTP_SUBMIT pid=${process.pid} token=${token} resolverFound=${!!resolver} pendingSize=${pendingForms.size} preSize=${preReceivedPayloads.size}\n`);
      if (resolver) {
        pendingForms.delete(token);
        pendingFormHtml.delete(token); // clean up served HTML
        try { resolver(JSON.parse(body)); } catch { resolver(body); }
        res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        res.end(`<!DOCTYPE html><html><head><title>Submitted</title><style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f0fdf4}div{text-align:center;color:#166534}</style></head><body><div><svg viewBox="0 0 24 24" width="48" height="48" fill="none" stroke="#16a34a" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg><h2>Form submitted!</h2><p>Claude is now generating your document.</p></div></body></html>`);
      } else {
        // wait_for_form_submit not called yet — buffer the payload so it can pick it up immediately when called.
        try { preReceivedPayloads.set(token, JSON.parse(body)); } catch { preReceivedPayloads.set(token, body); }
        res.writeHead(200, { "Content-Type": "text/plain" }); res.end("buffered");
      }
    });
  } else {
    res.writeHead(404); res.end();
  }
});
formHttpServer.listen(FORM_PORT, "127.0.0.1");
formHttpServer.on("error", () => { /* port in use — form falls back to copy-paste */ });

let currentToken = process.env.SEISMIC_API_TOKEN ?? "";
// True once a token was explicitly provided via set_token — disables the silent
// 401-triggered autoLogin() so it can never clobber a hand-picked token with a
// narrower-scoped one obtained from the default credential-flow login.
let tokenIsManual = false;

// Scopes requested by the credential-flow login (login tool + autoLogin refresh).
// Must include seismic.library.view/manage — search endpoints reject tokens
// without them ("Invalid or missing user claim"), even though generation
// endpoints are fine with just livedoc/library.
const LOGIN_SCOPE = "library livedoc seismic.library.view seismic.library.manage";

function authHeaders(): Record<string, string> {
  return {
    Authorization: `Bearer ${currentToken}`,
    "Content-Type": "application/json",
  };
}

async function autoLogin(): Promise<boolean> {
  if (!DEFAULT_AUTH_URI || !DEFAULT_AUTH_TENANT || !DEFAULT_CLIENT_ID || !DEFAULT_USERNAME || !DEFAULT_PASSWORD) return false;
  try {
    const body = new URLSearchParams({
      grant_type:    "client_credentials",
      client_id:     DEFAULT_CLIENT_ID,
      client_secret: DEFAULT_CLIENT_SECRET,
      username:      DEFAULT_USERNAME,
      password:      DEFAULT_PASSWORD,
      scope:         LOGIN_SCOPE,
    });
    const res = await fetch(`${DEFAULT_AUTH_URI}/tenants/${encodeURIComponent(DEFAULT_AUTH_TENANT)}/connect/token`, {
      method:  "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body:    body.toString(),
    });
    if (!res.ok) return false;
    const data = await res.json() as Record<string, unknown>;
    if (!data.access_token) return false;
    currentToken = data.access_token as string;
    tokenIsManual = false;
    return true;
  } catch {
    return false;
  }
}

async function apiFetch(
  path: string,
  options: RequestInit = {},
  _retry = true
): Promise<{ status: number; body: unknown }> {
  const url = `${BASE_URL}${path}`;
  const res = await fetch(url, {
    ...options,
    headers: { ...authHeaders(), ...(options.headers as Record<string, string> ?? {}) },
  });
  // Auto-refresh token on 401 when env credentials are available — but never when the
  // current token was explicitly set via set_token, so we don't silently replace a
  // hand-picked (possibly broader-scoped) token with a narrower credential-flow one.
  if (res.status === 401 && _retry && !tokenIsManual && DEFAULT_USERNAME && DEFAULT_PASSWORD) {
    const refreshed = await autoLogin();
    if (refreshed) return apiFetch(path, options, false);
  }
  const text = await res.text();
  let body: unknown;
  try {
    body = JSON.parse(text);
  } catch {
    body = text;
  }
  return { status: res.status, body };
}

const FORM_APP_BASE = process.env.FORM_APP_URL ?? "http://localhost:5173";
const FORM_API_BASE = process.env.FORM_API_URL ?? "http://localhost:3001";

function isComplex(resp: Record<string, unknown>): boolean {
  const vlData = (resp.variableListData ?? resp.VariableListData) as Array<Record<string, unknown>> | undefined;
  const imgInput = (resp.imageUploadContentInput ?? resp.ImageUploadContentInput) as Record<string, unknown> | undefined;
  const manualSelect = (resp.manualSelectContentInput ?? resp.ManualSelectContentInput) as Record<string, unknown> | undefined;
  const adhoc = (resp.adhocInputs ?? resp.AdhocInputs) as Array<unknown> | undefined;
  return !!(
    vlData?.some(v => v.dataSourceId ?? v.DataSourceId) ||
    (imgInput?.imageUploadContentItems as Array<unknown> | undefined)?.length ||
    (manualSelect?.manualSelectContentItems as Array<unknown> | undefined)?.length ||
    (adhoc?.length ?? 0) > 8
  );
}

// ── Tool definitions ────────────────────────────────────────────────────────

const FORM_RESOURCE_URI = "ui://livedoc/form";

// Tracks the URL of the most recently generated form for the App panel to fetch.
let latestFormUrl: string | null = null;

// Builds the MCP App panel shell HTML with the ext-apps bundle inlined.
// The bundle exposes globalThis.__McpApp.App so no bare-specifier import is needed.
function buildShellHtml(extAppsBundle: string): string {
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  html,body{margin:0;padding:0;width:100%;background:#fff2e0}
  #loading{display:flex;flex-direction:column;align-items:center;justify-content:center;
    min-height:120px;gap:10px;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
  #stage{font-size:18px;font-weight:700;color:#b45309}
  #detail{font-size:12px;color:#78350f;padding:2px 16px;text-align:center;
    word-break:break-all;max-width:380px;white-space:pre-wrap}
  #form-frame{display:none;width:100%;height:800px;border:0}
</style>
</head>
<body>
<div id="loading">
  <span id="stage">LIVEDOC — Initializing</span>
  <span id="detail">bundle loaded, constructing App…</span>
</div>
<iframe id="form-frame"></iframe>
<script>
${extAppsBundle}
</script>
<script>
  const st = (s, d) => {
    document.getElementById("stage").textContent = "LIVEDOC — " + s;
    if (d !== undefined) document.getElementById("detail").textContent = d;
  };
  const App = (globalThis.__McpApp || {}).App;
  if (!App) {
    st("BUNDLE ERROR", "globalThis.__McpApp.App not found after bundle");
  } else {
    const app = new App({ name: "livedoc-form", version: "1.0.0" }, {});
    const log = (m) => app.callServerTool({ name: "log_debug_message", arguments: { msg: m } }).catch(() => {});
    // ontoolresult carries structuredContent directly from get_livedoc_inputs.
    app.ontoolresult = async (event) => {
      log("ontoolresult-fired");
      try {
        const html = event?.structuredContent?.formHtml;
        const token = event?.structuredContent?.formToken;
        log("token:" + (token || "null") + " html-len:" + (html?.length || 0));
        if (html && html.length > 100) {
          const frame = document.getElementById("form-frame");
          frame.srcdoc = html;
          frame.style.display = "block";
          document.getElementById("loading").style.display = "none";
          st("form loaded", "");
          // Request a large panel height so the form is usable.
          app.sendSizeChanged({ width: 520, height: 800 });
        } else {
          st("NO FORM HTML", "structuredContent.formHtml missing (len=" + (html?.length || 0) + ")");
        }
      } catch(e) {
        st("tool-result error", String(e));
        log("toolresult-error:" + e.message);
      }
    };
    app.ontoolinput = async () => {
      log("ontoolinput-fired");
    };
    // Relay form submissions from the srcdoc iframe to the MCP server via callServerTool.
    // The MCP tool handler makes a server-side HTTP POST to port 3099 (no browser CSP applies).
    window.addEventListener("message", async (e) => {
      if (!e.data || e.data.type !== "livedoc-submit") return;
      const { token, payload } = e.data;
      log("form-submit-relay token=" + token);
      try {
        await app.callServerTool({ name: "receive_form_submission", arguments: { token, payload } });
        log("form-submit-relay ok");
      } catch(err) {
        log("form-submit-relay error:" + err.message);
      }
    });
    app.connect()
      .then(() => { st("connected — waiting for form", ""); log("connected"); })
      .catch(e => { st("CONNECT FAILED", String(e)); });
  }
</script>
</body>
</html>`;
}

// ── Form HTML builder ───────────────────────────────────────────────────────

function esc(s: string): string {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function gf(o: Record<string, unknown>, key: string): unknown {
  return o[key] ?? o[key[0].toUpperCase() + key.slice(1)];
}

function typeLabel(type: string): string {
  switch (type.toUpperCase()) {
    case "STRING":  return " (string)";
    case "INTEGER": return " (integer)";
    case "FLOAT":   return " (float)";
    case "DATE":    return " (date)";
    case "BOOL": case "BOOLEAN": return " (boolean)";
    default: return "";
  }
}

function scalarInput(name: string, type: string, scope: string, vlName?: string): string {
  const t = type.toUpperCase();
  const id = "f-" + name.replace(/[^a-zA-Z0-9]/g, "_");
  const vlAttr = vlName ? ` data-vl-name="${esc(vlName)}"` : "";
  const label = esc(name) + typeLabel(t);
  if (t === "BOOL" || t === "BOOLEAN") {
    return `<div class="bool-field"><input type="checkbox" id="${id}" data-scope="${scope}" data-field-name="${esc(name)}" data-field-type="${esc(t)}"${vlAttr}><label for="${id}">${label}</label></div>`;
  }
  const itype = t === "DATE" ? "date" : (t === "INTEGER" || t === "FLOAT") ? "number" : "text";
  const step = t === "FLOAT" ? ` step="any"` : t === "INTEGER" ? ` step="1"` : "";
  return `<div class="fw"><label class="fl" for="${id}">${label}</label><input type="${itype}"${step} id="${id}" class="fi" data-scope="${scope}" data-field-name="${esc(name)}" data-field-type="${esc(t)}"${vlAttr} placeholder="${esc(name)}"></div>`;
}

function tableInput(name: string, columns: Array<Record<string, unknown>>, scope: string, vlName?: string): string {
  const tid = "tbl-" + name.replace(/[^a-zA-Z0-9]/g, "_");
  const colDefs = columns.map(c => ({ name: String(gf(c, "name") ?? ""), type: String(gf(c, "type") ?? "STRING") }));
  const colsAttr = esc(JSON.stringify(colDefs));
  const vlAttr = vlName ? ` data-vl-name="${esc(vlName)}"` : "";
  const ths = colDefs.map(c => `<th>${esc(c.name)}</th>`).join("") + `<th style="width:32px"></th>`;
  return `<div class="tbl-wrap"><div class="sl">${esc(name)}</div><table class="dt" data-table-id="${tid}" data-table-scope="${scope}" data-table-name="${esc(name)}"${vlAttr} data-cols="${colsAttr}"><thead><tr>${ths}</tr></thead><tbody id="${tid}-body"></tbody></table><button class="add-btn" onclick="addRow('${tid}')">+ Add row</button></div>`;
}

function contentTypeIconSvg(contentType: string): string {
  const t = (contentType ?? "").toLowerCase();
  const svg = (body: string) => `<svg class="ct-icon" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">${body}</svg>`;
  if (t === "group") {
    // Folder
    return svg(`<path fill="#f59e0b" d="M1 3.5A1.5 1.5 0 012.5 2H6l1.5 2H13.5A1.5 1.5 0 0115 5.5v7A1.5 1.5 0 0113.5 14h-11A1.5 1.5 0 011 12.5z"/>`);
  }
  if (t === "section") {
    // Stacked layers
    return svg(`<path fill="#6b7280" d="M8.235 1.56a.5.5 0 00-.47 0l-7.5 4 7.5 4 7.5-4-7.5-4zm-7.13 8.96l7 3.734 7-3.734-1-.534L8 13.197 1.895 10.52l-1-.535-.789.535z"/>`);
  }
  if (t === "resourcepdf" || t === "pdf" || t === "resourcepdfpage") {
    // Document (red)
    return svg(`<path fill="#dc2626" d="M9.5 0H4a2 2 0 00-2 2v12a2 2 0 002 2h8a2 2 0 002-2V4.5L9.5 0zm0 1.5V4a.5.5 0 00.5.5H13V14a1 1 0 01-1 1H4a1 1 0 01-1-1V2a1 1 0 011-1h5.5z"/><path fill="#dc2626" d="M4.5 8h7v1h-7zm0 2h7v1h-7zm0 2h4v1h-4z"/>`);
  }
  // Default: presentation/slides
  return svg(`<path fill="#3b82f6" d="M0 2.5A1.5 1.5 0 011.5 1h13A1.5 1.5 0 0116 2.5v9A1.5 1.5 0 0114.5 13H9l.5.5H11a.5.5 0 010 1H5a.5.5 0 010-1h1.5l.5-.5H1.5A1.5 1.5 0 010 11.5zm1.5-.5a.5.5 0 00-.5.5v9a.5.5 0 00.5.5h13a.5.5 0 00.5-.5v-9a.5.5 0 00-.5-.5z"/>`);
}

function buildFormHtml(
  templateName: string,
  adhocInputs: Array<Record<string, unknown>>,
  varListInputs: Array<Record<string, unknown>>,
  manualSelect: Record<string, unknown> | undefined,
  forms: Array<Record<string, unknown>>,
  teamSiteId: string,
  versionId: string,
  token: string
): string {
  const isTable = (i: Record<string, unknown>) => ((gf(i, "columns") as unknown[] | undefined)?.length ?? 0) > 0;
  const scalarAdhoc = adhocInputs.filter(i => !isTable(i));
  const tableAdhoc  = adhocInputs.filter(i => isTable(i));

  const scalarGrid = scalarAdhoc.length
    ? `<div class="grid">${scalarAdhoc.map(i => scalarInput(String(gf(i, "name") ?? ""), String(gf(i, "type") ?? "STRING"), "adhoc")).join("")}</div>`
    : "";

  const tableHtml = tableAdhoc.map(i =>
    tableInput(String(gf(i, "name") ?? ""), (gf(i, "columns") as Array<Record<string, unknown>>) ?? [], "adhoc")
  ).join("");

  const vlHtml = varListInputs.map(vl => {
    const vlName   = String(gf(vl, "variableListName") ?? "");
    const dsName   = String(gf(vl, "dataSourceName") ?? gf(vl, "dataSourceId") ?? "");
    const inputs   = (gf(vl, "variableInputs") as Array<Record<string, unknown>>) ?? [];
    const scVl     = inputs.filter(i => !isTable(i));
    const tVl      = inputs.filter(i => isTable(i));
    const dsSpan   = dsName ? ` <span class="badge">${esc(dsName)}</span>` : "";
    const scGrid   = scVl.length ? `<div class="grid">${scVl.map(i => scalarInput(String(gf(i, "name") ?? ""), String(gf(i, "type") ?? "STRING"), "vl", vlName)).join("")}</div>` : "";
    const tblParts = tVl.map(i => tableInput(String(gf(i, "name") ?? ""), (gf(i, "columns") as Array<Record<string, unknown>>) ?? [], "vl", vlName)).join("");
    return `<div class="section"><div class="sl">Variable list — ${esc(vlName)}${dsSpan}</div>${scGrid}${tblParts}</div>`;
  }).join("");

  // Manual select groups/sections (plain include/exclude) and external content (one or more
  // resolved candidates picked via checkboxes — candidates are pre-attached by handleGetInputs).
  let msHtml = "";
  if (manualSelect) {
    const items = (gf(manualSelect, "manualSelectContentItems") as Array<Record<string, unknown>>) ?? [];
    const groups = items.filter(i => ["Group", "Section"].includes(String(gf(i, "contentType") ?? "")));
    const external = items.filter(i => !["Group", "Section"].includes(String(gf(i, "contentType") ?? "")));

    const hasGroupImages = groups.some(g => !!gf(g, "imageUrl"));
    const checks = groups.map(g => {
      const gId = esc(String(gf(g, "id") ?? ""));
      const gName = esc(String(gf(g, "name") ?? ""));
      const gType = String(gf(g, "contentType") ?? "Group");
      const gTypeEsc = esc(gType);
      const inc = gf(g, "isInclude") !== false ? " checked" : "";
      const oi = Number(gf(g, "orderIndex") ?? 0);
      const imgUrl = String(gf(g, "imageUrl") ?? "");
      const labelEl = `<label class="grp"><input type="checkbox"${inc} data-group-id="${gId}" data-group-name="${gName}" data-order-index="${oi}" data-content-type="${gTypeEsc}">${contentTypeIconSvg(gType)}<span>${gName}</span></label>`;
      if (imgUrl) {
        return `<div class="grp-card"><img class="grp-thumb" src="${esc(imgUrl)}" alt="${gName}" loading="lazy" crossorigin="anonymous" onerror="this.style.display='none'">${labelEl}</div>`;
      }
      return labelEl;
    }).join("");
    const listClass = hasGroupImages ? "grp-list grp-list-cards" : "grp-list";
    const groupsHtml = groups.length
      ? `<div class="section"><div class="sl">Content selection</div><div class="${listClass}">${checks}</div></div>`
      : "";

    const externalHtml = external.map(item => {
      const iId = esc(String(gf(item, "id") ?? ""));
      const iName = esc(String(gf(item, "name") ?? ""));
      const iContentType = String(gf(item, "contentType") ?? "");
      const iIcon = contentTypeIconSvg(iContentType);
      const oi = Number(gf(item, "orderIndex") ?? 0);
      const candidates = (gf(item, "candidates") as Array<Record<string, unknown>> | undefined) ?? [];
      if (!candidates.length) {
        return `<div class="ext-item"><div class="sl" style="margin-bottom:6px">${iIcon}${iName}</div><span class="badge" style="background:#fde8e8;color:#c00">No matching content found</span></div>`;
      }
      const totalCount = Number(gf(item, "candidatesTotalCount") ?? candidates.length);
      const truncatedNote = totalCount > candidates.length
        ? ` <span class="badge">Showing ${candidates.length} of ${totalCount} matches — refine the template's content filter if you need a different one</span>`
        : "";
      // Multiple documents can be attached to the same slot, so each candidate is its own
      // checkbox rather than a single-select dropdown — checking N boxes submits N items
      // that all share this slot's id/name but carry different resolved versionId/format.
      const candidateChecks = candidates.map((c, i) => {
        const val = esc(JSON.stringify({ versionId: gf(c, "versionId"), sourceBlobId: gf(c, "sourceBlobId"), format: gf(c, "format") }));
        const cFormat = String(gf(c, "format") ?? "");
        const cIcon = contentTypeIconSvg(cFormat.toLowerCase() === "pdf" ? "resourcepdf" : "liveslide");
        const label = esc(`${String(gf(c, "title") ?? "")} (${cFormat})`);
        const checkedAttr = i === 0 ? " checked" : "";
        return `<label class="grp"><input type="checkbox"${checkedAttr} data-external-candidate="${iId}" data-external-name="${iName}" data-order-index="${oi}" value='${val}'>${cIcon}<span>${label}</span></label>`;
      }).join("");
      return `<div class="ext-item"><div class="sl" style="margin-bottom:6px">${iIcon}${iName}${truncatedNote}</div><div class="grp-list">${candidateChecks}</div></div>`;
    }).join("");

    msHtml = groupsHtml + (externalHtml ? `<div class="section"><div class="sl">External content</div>${externalHtml}</div>` : "");
  }

  // Group forms by unique name → { formName: [{outputs}, ...] }
  type FormCfg = { outputs: Array<{ format: unknown }> };
  const formsByName = new Map<string, FormCfg[]>();
  for (const f of forms) {
    const name = String(gf(f, "name") ?? "");
    if (!formsByName.has(name)) formsByName.set(name, []);
    formsByName.get(name)!.push({
      outputs: ((gf(f, "outputs") as Array<Record<string, unknown>>) ?? []).map(o => ({ format: gf(o, "format") })),
    });
  }
  const uniqueFormNames = Array.from(formsByName.keys());
  const multiForm = uniqueFormNames.length > 1;

  // Initial state: first form name, first output combo
  const firstFormName = uniqueFormNames[0] ?? "";
  const firstFormCfgs = formsByName.get(firstFormName) ?? [];
  const initOutputs = JSON.stringify(firstFormCfgs[0]?.outputs ?? []);

  // Form selector (only when >1 distinct form name)
  const formSelHtml = multiForm
    ? `<div class="fmt-row"><div class="sl" style="margin-bottom:10px">Select form</div><div id="form-sel">${
        uniqueFormNames.map((n, i) =>
          `<button class="fmt${i === 0 ? " active" : ""}" data-form-name="${esc(n)}" onclick="selForm(this)">${esc(n)}</button>`
        ).join("")
      }</div></div>`
    : "";

  // Initial output format buttons (format codes, not form names)
  const fmtBtnsHtml = firstFormCfgs.map((cfg, i) => {
    const label = esc(cfg.outputs.map(o => String(o.format)).join(" + ") || "Default");
    return `<button class="fmt${i === 0 ? " active" : ""}" data-outputs="${esc(JSON.stringify(cfg.outputs))}" onclick="selFmt(this)">${label}</button>`;
  }).join("");

  // JS config: all form configs keyed by name
  const formConfigsJs = JSON.stringify(
    Object.fromEntries(Array.from(formsByName.entries()))
  );

  const css = `*{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
body{background:#fff;padding:20px;font-size:14px;color:#1d1d1f}
.title{font-size:17px;font-weight:700;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px 16px;margin-bottom:20px;align-items:end}
.fw{display:flex;flex-direction:column;gap:5px}
.fl{font-size:12px;font-weight:600;color:#555}
.fi{width:100%;padding:7px 10px;border:1px solid #d0d0d0;border-radius:7px;font-size:13px;outline:none;background:#fff}
.fi:focus{border-color:#0066cc}
.bool-field{display:flex;align-items:center;gap:8px;padding-bottom:10px}
.bool-field input{width:16px;height:16px;cursor:pointer;flex-shrink:0}
.bool-field label{font-size:12px;font-weight:600;color:#555;cursor:pointer}
.section{margin-bottom:20px}
.sl{font-size:13px;font-weight:700;color:#444;margin-bottom:10px}
.badge{display:inline-block;padding:1px 7px;border-radius:100px;font-size:11px;font-weight:600;background:#e8f0fe;color:#0066cc;margin-left:6px}
.tbl-wrap{margin-bottom:20px}
.dt{width:100%;border-collapse:collapse;font-size:13px}
.dt th{font-size:12px;font-weight:600;color:#666;padding:6px 8px;text-align:left;border-bottom:1px solid #e5e5e5;background:#fafafa}
.dt td{padding:4px 6px;border-bottom:1px solid #f5f5f5}
.dt td input[type=text],.dt td input[type=number],.dt td input[type=date]{width:100%;padding:5px 7px;border:1px solid #d0d0d0;border-radius:5px;font-size:12px;outline:none}
.dt td input[type=text]:focus,.dt td input[type=number]:focus,.dt td input[type=date]:focus{border-color:#0066cc}
.dt td input[type=checkbox]{width:16px;height:16px;cursor:pointer}
.del{background:none;border:none;cursor:pointer;color:#ccc;font-size:15px;padding:2px 5px}
.del:hover{color:#c00}
.add-btn{font-size:12px;color:#0066cc;background:none;border:1px dashed #0066cc;border-radius:6px;padding:5px 14px;cursor:pointer;margin-top:6px}
.add-btn:hover{background:#e8f0fe}
.fmt-row{margin-bottom:20px}
.fmt{padding:7px 18px;border-radius:20px;border:none;background:#f0f0f0;color:#444;font-size:13px;font-weight:500;cursor:pointer;margin-right:8px;transition:background .15s}
.fmt.active{background:#0066cc;color:#fff}
.sub{display:inline-flex;align-items:center;gap:6px;padding:9px 22px;background:#0066cc;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer}
.sub:hover{background:#0055b3}
.grp-list{display:flex;flex-direction:column;gap:8px}
.grp{display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px}
.grp input{width:16px;height:16px}
.ext-item{margin-bottom:14px;padding:10px 12px;border:1px solid #e5e5e5;border-radius:8px}
.ct-icon{width:14px;height:14px;display:inline-block;vertical-align:middle;margin-right:5px;flex-shrink:0}
.grp-list-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px}
.grp-card{border:1px solid #e5e5e5;border-radius:8px;overflow:hidden;cursor:pointer;transition:box-shadow .15s}
.grp-card:hover{box-shadow:0 2px 8px rgba(0,0,0,.1)}
.grp-card:has(input:checked){border-color:#0066cc;box-shadow:0 0 0 2px rgba(0,102,204,.15)}
.grp-thumb{width:100%;height:80px;object-fit:cover;display:block;background:#f0f2f5}
.grp-card .grp{padding:8px 10px}`;

  const js = `var sel=${initOutputs};
var tsid=${JSON.stringify(teamSiteId)};
var vid=${JSON.stringify(versionId)};
var formToken=${JSON.stringify(token)};
var formCfgs=${formConfigsJs};
function selFmt(b){
  sel=JSON.parse(b.dataset.outputs);
  document.querySelectorAll('#fmt-btns .fmt').forEach(function(x){x.classList.remove('active')});
  b.classList.add('active');
}
function selForm(b){
  document.querySelectorAll('#form-sel .fmt').forEach(function(x){x.classList.remove('active')});
  b.classList.add('active');
  var cfgs=formCfgs[b.dataset.formName]||[];
  var div=document.getElementById('fmt-btns');
  div.innerHTML='';
  cfgs.forEach(function(cfg,i){
    var btn=document.createElement('button');
    btn.className='fmt'+(i===0?' active':'');
    btn.dataset.outputs=JSON.stringify(cfg.outputs);
    btn.textContent=cfg.outputs.map(function(o){return o.format;}).join(' + ')||'Default';
    btn.onclick=function(){selFmt(this);};
    div.appendChild(btn);
  });
  if(cfgs.length)sel=cfgs[0].outputs;
}
function addRow(tid){
  var tbl=document.querySelector('[data-table-id="'+tid+'"]');
  if(!tbl)return;
  var cols=JSON.parse(tbl.dataset.cols||'[]');
  var tb=document.getElementById(tid+'-body');
  if(!tb)return;
  var tr=document.createElement('tr');
  cols.forEach(function(c){
    var td=document.createElement('td');
    var t=(c.type||'').toUpperCase();
    var el;
    if(t==='BOOL'||t==='BOOLEAN'){el=document.createElement('input');el.type='checkbox';}
    else if(t==='DATE'){el=document.createElement('input');el.type='date';}
    else{el=document.createElement('input');el.type=(t==='INTEGER'||t==='FLOAT')?'number':'text';if(t==='INTEGER')el.step='1';if(t==='FLOAT')el.step='any';}
    el.dataset.colType=t;td.appendChild(el);tr.appendChild(td);
  });
  var dtd=document.createElement('td');
  var db=document.createElement('button');
  db.className='del';db.innerHTML='&#128465;';
  db.onclick=function(){this.closest('tr').remove();};
  dtd.appendChild(db);tr.appendChild(dtd);tb.appendChild(tr);
}
function tblVal(tbl){
  var cols=JSON.parse(tbl.dataset.cols||'[]');
  var rows=Array.from(tbl.querySelectorAll('tbody tr')).map(function(tr){
    return Array.from(tr.querySelectorAll('[data-col-type]')).map(function(c){
      var t=(c.dataset.colType||'').toUpperCase();
      if(t==='BOOL'||t==='BOOLEAN')return c.checked;
      if(t==='INTEGER')return parseInt(c.value)||0;
      if(t==='FLOAT')return parseFloat(c.value)||0;
      return c.value;
    });
  });
  return{columns:cols.map(function(c){return c.name;}),rows:rows};
}
function scVal(el){
  var t=(el.dataset.fieldType||'').toUpperCase();
  if(t==='BOOL'||t==='BOOLEAN')return el.checked;
  if(t==='INTEGER')return parseInt(el.value)||0;
  if(t==='FLOAT')return parseFloat(el.value)||0;
  return el.value;
}
function submit(){
  var adhoc=[];
  document.querySelectorAll('[data-scope="adhoc"]').forEach(function(el){adhoc.push({name:el.dataset.fieldName,value:scVal(el)});});
  document.querySelectorAll('.dt[data-table-scope="adhoc"]').forEach(function(t){adhoc.push({name:t.dataset.tableName,value:tblVal(t)});});
  var vlmap={};
  document.querySelectorAll('[data-scope="vl"]').forEach(function(el){var v=el.dataset.vlName;if(!vlmap[v])vlmap[v]=[];vlmap[v].push({name:el.dataset.fieldName,value:scVal(el)});});
  document.querySelectorAll('.dt[data-table-scope="vl"]').forEach(function(t){var v=t.dataset.vlName;if(!vlmap[v])vlmap[v]=[];vlmap[v].push({name:t.dataset.tableName,value:tblVal(t)});});
  // Skip variable list entries where every input is empty (empty string or table with no rows).
  var vld=Object.keys(vlmap).filter(function(k){
    return vlmap[k].some(function(inp){
      var v=inp.value;
      if(v===null||v===undefined||v==='')return false;
      if(typeof v==='object'&&Array.isArray(v.rows))return v.rows.length>0;
      return true;
    });
  }).map(function(k){return{variableListName:k,variableInputs:vlmap[k]};});
  var ms=[];
  document.querySelectorAll('[data-group-id]').forEach(function(cb){ms.push({id:cb.dataset.groupId,name:cb.dataset.groupName,contentType:cb.dataset.contentType||'Group',isInclude:cb.checked,orderIndex:parseInt(cb.dataset.orderIndex)||0});});
  // Multiple documents can be attached to the same external-content slot: group checkboxes by
  // slot id, then emit one manualSelectContentItem PER CHECKED candidate (all sharing that
  // slot's id/name), or a single isInclude:false item if none are checked.
  var extGroups={};
  document.querySelectorAll('[data-external-candidate]').forEach(function(cb){
    var id=cb.dataset.externalCandidate;
    if(!extGroups[id])extGroups[id]={name:cb.dataset.externalName,orderIndex:parseInt(cb.dataset.orderIndex)||0,checked:[]};
    if(cb.checked)extGroups[id].checked.push(JSON.parse(cb.value));
  });
  Object.keys(extGroups).forEach(function(id){
    var g=extGroups[id];
    if(g.checked.length){
      g.checked.forEach(function(chosen){
        var item={id:id,name:g.name,isInclude:true,orderIndex:g.orderIndex,versionId:chosen.versionId,contentType:(chosen.format||'').toUpperCase()==='PDF'?'ResourcePDF':'LiveSlide'};
        if(chosen.sourceBlobId)item.sourceBlobId=chosen.sourceBlobId;
        ms.push(item);
      });
    }else{
      ms.push({id:id,name:g.name,contentType:'LiveSlide',isInclude:false,orderIndex:g.orderIndex});
    }
  });
  var p={teamSiteId:tsid,libraryContentVersionId:vid,adHocInputs:adhoc,outputs:sel};
  if(vld.length)p.variableListData=vld;
  if(ms.length)p.manualSelectContentInput={manualSelectContentItems:ms};
  var msg=JSON.stringify(p);
  var btn=document.getElementById('sub-btn');
  btn.disabled=true;btn.textContent='Submitting…';
  // Send via postMessage to parent App panel shell (avoids browser fetch CSP restrictions).
  // The shell relays it to the MCP server via callServerTool.
  window.parent.postMessage({type:'livedoc-submit',token:formToken,payload:msg},'*');
  // Show success immediately — shell will signal back on error.
  btn.style.display='none';
  document.getElementById('done-msg').style.display='block';
}
function copyPayload(){
  var t=document.getElementById('payload-text');
  t.select();
  try{navigator.clipboard.writeText(t.value).then(function(){document.getElementById('copy-btn').textContent='Copied!';}).catch(function(){document.execCommand('copy');document.getElementById('copy-btn').textContent='Copied!';});}
  catch(e){try{document.execCommand('copy');document.getElementById('copy-btn').textContent='Copied!';}catch(e2){}}
}`;

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${css}</style></head><body>
<div class="title">${esc(templateName)} — inputs</div>
${scalarGrid}${tableHtml}${vlHtml}${msHtml}
${formSelHtml}
<div class="fmt-row"><div class="sl" style="margin-bottom:10px">Output format</div><div id="fmt-btns">${fmtBtnsHtml}</div></div>
<button id="sub-btn" class="sub" onclick="submit()">&#9654; Submit generation</button>
<div id="done-msg" style="display:none;margin-top:16px;padding:14px;background:#f0faf0;border:1.5px solid #b2dfb2;border-radius:8px;color:#2e7d32;font-weight:600;font-size:14px">&#10003; Submitted! Generation starting…</div>
<div id="payload-box" style="display:none;margin-top:16px;padding:14px;background:#f0f7ff;border:1.5px solid #90b8e8;border-radius:8px">
  <div style="font-weight:600;font-size:13px;margin-bottom:8px;color:#0055aa">Copy this payload and paste it into the chat:</div>
  <textarea id="payload-text" readonly style="width:100%;height:72px;font-size:11px;font-family:monospace;border:1px solid #b0c8e8;border-radius:4px;padding:6px;box-sizing:border-box;resize:none;background:#fff"></textarea>
  <button id="copy-btn" onclick="copyPayload()" style="margin-top:8px;padding:7px 20px;background:#0066cc;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600">&#128203; Copy to clipboard</button>
</div>
<div id="paste-hint" style="display:none;margin-top:12px;padding:12px 14px;background:#fff8e1;border:1.5px solid #f9a825;border-radius:8px;font-size:13px;color:#5d4037">
  <b>Auto-submit unavailable in this view.</b><br>The payload has been copied to your clipboard. Paste it into the chat and Claude will continue from there.
</div>
<script>${js}<\/script>
</body></html>`;
}

// ── Tool handlers ───────────────────────────────────────────────────────────

async function handleSearchTemplates(args: {
  searchText?: string;
  page_size?: number;
}) {
  const size = Math.min(args.page_size ?? 10, 50);
  const body = {
    searchText: args.searchText ?? "",
    allowPptx: true,
    includeLiveDoc: true,
    allowPdf: false,
    page: { size, from: 0 },
    orderBy: [{ attr: "modifiedDate", direction: "DESC" }],
  };
  const result = await apiFetch("/v3/contents", {
    method: "POST",
    body: JSON.stringify(body),
  });
  if (result.status !== 200) {
    const detail = result.body as Record<string, unknown> | undefined;
    const isUserClaimError = typeof detail === "object" && String(detail?.Message ?? "").includes("user claim");
    if (isUserClaimError) {
      return {
        error: "search_requires_user_token",
        message: "Template search requires a user-context token. The current token is a service account token without user identity claims. To enable search: set SEISMIC_API_TOKEN in claude_desktop_config.json to a user token (obtain one from the Seismic dev portal or browser DevTools). If you already know your template's teamSiteId and libraryContentVersionId, call get_livedoc_inputs directly — authentication for generation is not affected.",
      };
    }
    return { error: `Search failed (HTTP ${result.status})`, detail: result.body };
  }
  const data = result.body as {
    totalCount: number;
    documents: Array<{
      contentId: string;
      contentVersionId: string;
      title: string;
      description: string | null;
      format: string;
      teamsite: string;
      modifiedDate: string;
    }>;
  };
  return {
    totalCount: data.totalCount,
    results: data.documents.map((d) => ({
      title: d.title,
      format: d.format,
      contentVersionId: d.contentVersionId,
      teamSiteId: d.teamsite,
      modifiedDate: d.modifiedDate,
      description: d.description,
    })),
  };
}

async function handleSearchContent(args: {
  query: string;
  contentType?: string;
  page_size?: number;
}) {
  const size = Math.min(args.page_size ?? 10, 50);
  // Map contentType to /v3/contents flags. allowPptx requires at least one of
  // includeStandardPptx/includeLiveDoc also true, or the API rejects the request.
  const ct = args.contentType ?? "";
  const isSlideType = !ct || ["ExternalSlides", "LiveSlide", "ExternalStaticSlides"].includes(ct);
  const isLiveDocType = !ct || ct === "LiveDoc";
  const allowPptx = isSlideType || isLiveDocType;
  const includeStandardPptx = isSlideType;
  const includeLiveDoc = isSlideType || isLiveDocType;
  const allowPdf = !ct || ct === "PDF";
  const body = {
    searchText: args.query,
    allowPptx,
    includeStandardPptx,
    includeLiveDoc,
    allowPdf,
    page: { size, from: 0 },
    orderBy: [{ attr: "modifiedDate", direction: "DESC" }],
  };
  const result = await apiFetch("/v3/contents", {
    method: "POST",
    body: JSON.stringify(body),
  });
  if (result.status !== 200) {
    return { error: `Search failed (HTTP ${result.status})`, detail: result.body };
  }
  const data = result.body as {
    totalCount: number;
    documents: Array<{
      contentId: string;
      contentVersionId: string;
      title: string;
      format: string;
      teamsite: string;
      modifiedDate: string;
      sourceBlobId?: string;
    }>;
  };
  return {
    totalCount: data.totalCount,
    results: data.documents.map((d) => ({
      id: d.contentVersionId,
      name: d.title,
      format: d.format,
      contentVersionId: d.contentVersionId,
      sourceBlobId: d.sourceBlobId,
      modifiedDate: d.modifiedDate,
    })),
  };
}

// "Group"/"Section" are the only manualSelectContentItem types that are already fully valid
// as returned — every other contentType needs real content resolved via search before submission.
// This is a denylist rather than an allowlist because the GET side's vocabulary doesn't match
// the submission-side ManualSelectContentType enum 1:1 (e.g. GET can return "ExternalSlides",
// which isn't even a valid value to submit — it must be resolved then re-mapped to "LiveSlide"
// or "ResourcePDF" depending on the chosen candidate's format).
function needsContentResolution(contentType: string): boolean {
  return contentType !== "" && contentType !== "Group" && contentType !== "Section";
}

// C# bool property names here don't follow simple camelCase (AllowPDF, IncludeStandardPPTX),
// so check several literal casings rather than relying on gf()'s single-fallback capitalization.
function boolField(item: Record<string, unknown>, ...keys: string[]): boolean {
  for (const k of keys) {
    if (typeof item[k] === "boolean") return item[k] as boolean;
  }
  return false;
}

// Resolves real content candidates for one manualSelectContentItem. Prefers the item's own
// filter/format flags (the template author's actual search criteria, e.g. Filter: [{propertyName:
// "ContentName", operator: "CT", value: "sp3"}]) over a generic name-based guess — those flags
// are what get_livedoc_inputs actually returns on ExternalSlideContent items.
const CANDIDATE_PAGE_SIZE = 10;

async function resolveManualSelectCandidates(item: Record<string, unknown>): Promise<{ candidates: Array<Record<string, unknown>>; totalCount: number }> {
  const name = String(gf(item, "name") ?? "");
  const contentType = String(gf(item, "contentType") ?? "");
  const filter = (gf(item, "filter") as unknown[] | undefined) ?? [];
  const rawIsApplyAllFilter = gf(item, "isApplyAllFilter");
  const isApplyAllFilter = typeof rawIsApplyAllFilter === "boolean" ? rawIsApplyAllFilter : true;

  let allowPptx = boolField(item, "allowPptx", "AllowPptx");
  let includeStandardPptx = boolField(item, "includeStandardPptx", "IncludeStandardPPTX", "IncludeStandardPptx");
  let includeLiveDoc = boolField(item, "includeLiveDoc", "IncludeLiveDoc");
  let allowPdf = boolField(item, "allowPdf", "AllowPDF", "AllowPdf");

  // Fall back to a contentType-based guess only if the item carried no usable format flags at all.
  if (!allowPptx && !allowPdf) {
    const isSlideType = ["ExternalSlides", "LiveSlide", "ExternalStaticSlides"].includes(contentType);
    allowPptx = isSlideType;
    includeStandardPptx = isSlideType;
    includeLiveDoc = isSlideType;
    allowPdf = !isSlideType;
  }

  const body: Record<string, unknown> = {
    allowPptx,
    includeStandardPptx,
    includeLiveDoc,
    allowPdf,
    page: { size: CANDIDATE_PAGE_SIZE, from: 0 },
    orderBy: [{ attr: "modifiedDate", direction: "DESC" }],
  };
  if (filter.length > 0) {
    // The item's own filter is the template author's actual search criteria — combining it
    // with a searchText:name guess (name is just a display label, e.g. "sp3") over-constrains
    // the query and silently returns zero results, so filter and searchText are mutually exclusive here.
    body.filter = filter;
    body.isApplyAllFilter = isApplyAllFilter;
  } else {
    body.searchText = name;
  }

  const result = await apiFetch("/v3/contents", { method: "POST", body: JSON.stringify(body) });
  if (result.status !== 200) return { candidates: [], totalCount: 0 };
  const data = result.body as { documents?: Array<Record<string, unknown>>; totalCount?: number };
  const candidates = (data.documents ?? []).slice(0, CANDIDATE_PAGE_SIZE).map((d) => ({
    versionId: gf(d, "contentVersionId"),
    sourceBlobId: gf(d, "sourceBlobId"),
    title: gf(d, "title"),
    format: gf(d, "format"),
  }));
  return { candidates, totalCount: data.totalCount ?? candidates.length };
}

async function handleGetInputs(args: {
  teamSiteId: string;
  libraryContentVersionId: string;
}): Promise<{
  templateName: string;
  teamSiteId: string;
  libraryContentVersionId: string;
  adhocInputs: Array<Record<string, unknown>>;
  variableListData: Array<Record<string, unknown>>;
  manualSelectContentInput: Record<string, unknown> | undefined;
  forms: Array<Record<string, unknown>>;
  isComplex: boolean;
  hasImageUpload: boolean;
} | { error: string; detail: unknown }> {
  const result = await apiFetch(
    `/v3/teamsites/${args.teamSiteId}/livedocVersions/${args.libraryContentVersionId}`
  );
  if (result.status !== 200) {
    return { error: `Failed to get inputs (HTTP ${result.status})`, detail: result.body };
  }
  const raw = result.body as Record<string, unknown>;
  const templateName = String(raw.name ?? raw.Name ?? "Template");
  const adhocInputs = (raw.adhocInputs ?? raw.AdhocInputs) as Array<Record<string, unknown>> | undefined ?? [];
  const varListInputs = (raw.variableListData ?? raw.VariableListData) as Array<Record<string, unknown>> | undefined ?? [];
  const imageUpload = (raw.imageUploadContentInput ?? raw.ImageUploadContentInput) as Record<string, unknown> | undefined;
  const manualSelect = (raw.manualSelectContentInput ?? raw.ManualSelectContentInput) as Record<string, unknown> | undefined;
  const forms = (raw.forms ?? raw.Forms) as Array<Record<string, unknown>> ?? [];
  const complex = isComplex(raw);
  const hasImageUpload = !!((imageUpload as Record<string, unknown> | undefined)?.imageUploadContentItems as unknown[] | undefined)?.length;

  // Resolve external-content slots server-side instead of relying on the calling model to
  // remember a separate search_livedoc_content step — that step was repeatedly skipped in
  // practice, leaving the artifact with no real candidates to pick from.
  const msItems = (manualSelect ? gf(manualSelect, "manualSelectContentItems") : undefined) as Array<Record<string, unknown>> | undefined ?? [];
  await Promise.all(
    msItems.map(async (item) => {
      const contentType = String(gf(item, "contentType") ?? "");
      if (!needsContentResolution(contentType)) return;
      const resolved = await resolveManualSelectCandidates(item);
      item.candidates = resolved.candidates;
      item.candidatesTotalCount = resolved.totalCount;
    })
  );

  return {
    templateName,
    teamSiteId: args.teamSiteId,
    libraryContentVersionId: args.libraryContentVersionId,
    adhocInputs,
    variableListData: varListInputs,
    manualSelectContentInput: manualSelect,
    forms,
    isComplex: complex,
    hasImageUpload,
  };
}

async function handleSubmitGeneration(args: {
  teamSiteId: string;
  libraryContentVersionId: string;
  adHocInputs: Array<{ name: string; value: unknown }>;
  outputs: Array<{ format: string; name?: string; fileName?: string }>;
  variableListData?: Array<{
    variableListName: string;
    variableInputs: Array<{ name: string; value: unknown }>;
  }>;
  liveFormSellerTemplateId?: string;
  regionalFormat?: string;
  manualSelectContentInput?: {
    manualSelectContentItems: Array<{
      id: string;
      name?: string;
      contentType: string;
      versionId?: string;
      sourceBlobId?: string;
      pageNumber?: number;
      isInclude: boolean;
      orderIndex?: number;
    }>;
  };
}) {
  if (args.manualSelectContentInput) {
    const unresolved = args.manualSelectContentInput.manualSelectContentItems.filter(
      (item) =>
        item.isInclude !== false &&
        needsContentResolution(item.contentType) &&
        !item.versionId
    );
    if (unresolved.length > 0) {
      return {
        error: "WRONG TOOL — do not call submit_livedoc_generation directly when manualSelectContentInput has unresolved items.",
        detail: `Item(s) [${unresolved.map((i) => `"${i.name ?? i.id}"`).join(", ")}] are missing versionId. You must NOT resolve versionId yourself via search_livedoc_content or any other tool. The correct flow is: (1) get_livedoc_inputs opens an HTML form, (2) the USER fills in the form and clicks Submit, (3) you call wait_for_form_submit to receive the fully-resolved payload, (4) THEN call this tool with that exact payload. If you have not yet called wait_for_form_submit, call it now with the token from get_livedoc_inputs.`,
      };
    }
  }

  const reqBody: Record<string, unknown> = {
    adHocInputs: args.adHocInputs,
    outputs: args.outputs,
  };
  if (args.variableListData) reqBody.variableListData = args.variableListData;
  if (args.regionalFormat) reqBody.regionalFormat = args.regionalFormat;
  if (args.manualSelectContentInput) {
    reqBody.manualSelectContentInput = {
      manualSelectContentItems: args.manualSelectContentInput.manualSelectContentItems.map((item) => ({
        id: item.id,
        name: item.name,
        contentType: item.contentType,
        versionId: item.versionId,
        sourceBlobId: item.sourceBlobId,
        pageNumber: item.pageNumber,
        isInclude: item.isInclude,
        orderIndex: item.orderIndex,
      })),
    };
  }

  const qp = args.liveFormSellerTemplateId
    ? `?liveFormSellerTemplateId=${encodeURIComponent(args.liveFormSellerTemplateId)}`
    : "";

  const result = await apiFetch(
    `/v3/teamsites/${args.teamSiteId}/livedocVersions/${args.libraryContentVersionId}${qp}`,
    { method: "POST", body: JSON.stringify(reqBody) }
  );
  if (result.status !== 201 && result.status !== 200) {
    return { error: `Generation submission failed (HTTP ${result.status})`, detail: result.body };
  }
  const body = result.body as Record<string, unknown>;
  const generatedLivedocId =
    (body.generatedLivedocId ?? body.id ?? body.GeneratedLivedocId ?? body.Id) as string | undefined;
  return {
    generatedLivedocId,
    rawBody: body,
    message: "Generation job submitted. Call get_generation_status to poll for completion.",
  };
}

// Matches LiveDocGenStatusResp in app-livedoc-service (PublicAPIV3Controller.ResultStatus.cs):
// Queued=0, Generating=1, Completed=2, Failed=3. The API returns this as a number, not a string,
// so callers must map it before comparing against status names.
const STATUS_NAMES = ["Queued", "Generating", "Completed", "Failed"];

function statusName(raw: unknown): string {
  if (typeof raw === "number" && STATUS_NAMES[raw] !== undefined) {
    return STATUS_NAMES[raw];
  }
  if (typeof raw === "string" && STATUS_NAMES.includes(raw)) {
    return raw;
  }
  return String(raw);
}

async function handleGetStatus(args: { generatedLivedocId: string }) {
  const result = await apiFetch(`/v3/generatedLivedocs/${args.generatedLivedocId}`);
  if (result.status !== 200) {
    return { error: `Status check failed (HTTP ${result.status})`, detail: result.body };
  }
  const raw = result.body as Record<string, unknown>;
  const id = (raw.id ?? raw.Id ?? raw.generatedLivedocId ?? raw.GeneratedLivedocId) as string;
  const rawOutputs = (raw.outputs ?? raw.Outputs ?? []) as Array<Record<string, unknown>>;
  const outputs = rawOutputs.map((o) => ({
    id: (o.id ?? o.Id) as string,
    status: statusName(o.status ?? o.Status),
    format: (o.format ?? o.Format) as string,
    name: (o.name ?? o.Name) as string,
    fileName: (o.fileName ?? o.FileName) as string,
    errorString: (o.errorString ?? o.ErrorString ?? null) as string | null,
  }));
  const allDone = outputs.every((o) => o.status === "Completed" || o.status === "Failed");
  return {
    generatedLivedocId: id,
    allDone,
    outputs,
    hint: allDone
      ? "All outputs done. Call get_generation_download_url with each outputId."
      : "Still generating. Poll again in a few seconds.",
  };
}

async function handleGetDownloadUrl(args: {
  generatedLivedocId: string;
  outputId: string;
}) {
  const result = await apiFetch(
    `/v3/generatedLivedocs/${args.generatedLivedocId}/outputs/${args.outputId}/content?redirect=false`
  );
  if (result.status !== 200) {
    return { error: `Download URL fetch failed (HTTP ${result.status})`, detail: result.body };
  }
  return result.body;
}

function getDownloadsDir(): string {
  const dir = path.join(os.homedir(), "Downloads");
  try {
    fs.mkdirSync(dir, { recursive: true });
    return dir;
  } catch {
    return os.tmpdir();
  }
}

// Opens a local file with the OS-registered default application (e.g. double-click behavior).
function openWithDefaultApp(filePath: string) {
  const platform = process.platform;
  const quoted = `"${filePath}"`;
  const cmd =
    platform === "win32" ? `start "" ${quoted}` :
    platform === "darwin" ? `open ${quoted}` :
    `xdg-open ${quoted}`;
  exec(cmd, () => { /* best-effort; failures are non-fatal */ });
}

function uniqueFilePath(dir: string, fileName: string): string {
  const ext = path.extname(fileName);
  const base = path.basename(fileName, ext);
  let candidate = path.join(dir, fileName);
  let n = 1;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${base} (${n})${ext}`);
    n++;
  }
  return candidate;
}

async function handleDownloadGenerationOutput(args: {
  generatedLivedocId: string;
  outputId: string;
  autoOpen?: boolean;
}) {
  const dl = await handleGetDownloadUrl({ generatedLivedocId: args.generatedLivedocId, outputId: args.outputId });
  if (dl && typeof dl === "object" && "error" in (dl as object)) {
    return dl;
  }
  const body = dl as Record<string, unknown>;
  // DownloadLocationResp only ever contains `downloadUrl` — it never carries a fileName.
  // The real fileName (with extension) lives on get_generation_status's outputs[], so look it up there.
  const url = String(body.url ?? body.downloadUrl ?? body.Url ?? body.DownloadUrl ?? "");
  if (!url) {
    return { error: "No download URL returned for this output.", detail: body };
  }

  let fileName = `livedoc-${args.outputId}`;
  const status = await handleGetStatus({ generatedLivedocId: args.generatedLivedocId });
  if (status && typeof status === "object" && "outputs" in (status as object)) {
    const outputs = (status as { outputs: Array<{ id: string; format: string; fileName: string }> }).outputs;
    const match = outputs.find(
      (o) => o.id === args.outputId || (o.format ?? "").toLowerCase() === args.outputId.toLowerCase()
    );
    if (match) {
      fileName = match.fileName || `${fileName}.${(match.format ?? "").toLowerCase()}`;
    }
  }

  const fileRes = await fetch(url);
  if (!fileRes.ok) {
    return { error: `Failed to download file content (HTTP ${fileRes.status})` };
  }
  const buf = Buffer.from(await fileRes.arrayBuffer());

  const dir = getDownloadsDir();
  const filePath = uniqueFilePath(dir, fileName);
  fs.writeFileSync(filePath, buf);

  const autoOpen = args.autoOpen !== false;
  if (autoOpen) openWithDefaultApp(filePath);

  return {
    filePath,
    fileName,
    sizeBytes: buf.length,
    opened: autoOpen,
    message: autoOpen
      ? `Downloaded "${fileName}" (${buf.length} bytes) to ${filePath} and opened it with the default app.`
      : `Downloaded "${fileName}" (${buf.length} bytes) to ${filePath}.`,
  };
}

function generateToken(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

async function handleWaitForFormSubmit(args: { token: string }): Promise<unknown> {
  const dbg = (msg: string) => fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"), `[${new Date().toISOString()}] WAIT_SUBMIT pid=${process.pid} token=${args.token} ${msg}\n`);
  dbg(`called preSize=${preReceivedPayloads.size} pendingSize=${pendingForms.size}`);
  const submitFile = path.join(os.tmpdir(), `livedoc-submit-${args.token}.json`);
  const TIMEOUT_MS = 10 * 60 * 1000;
  return new Promise((resolve, reject) => {
    const done = (payload: unknown) => {
      clearTimeout(timer);
      clearInterval(poller);
      pendingForms.delete(args.token);
      dbg("resolved");
      resolve(payload);
    };
    const timer = setTimeout(() => {
      clearInterval(poller);
      pendingForms.delete(args.token);
      dbg("timeout");
      reject(new Error("Form submission timed out after 10 minutes."));
    }, TIMEOUT_MS);
    // In-process resolver (same process as receive_form_submission).
    pendingForms.set(args.token, (payload) => done(payload));
    // Cross-process file poller (different process wrote the file).
    const poller = setInterval(() => {
      if (fs.existsSync(submitFile)) {
        try {
          const raw = fs.readFileSync(submitFile, "utf-8");
          fs.unlinkSync(submitFile);
          dbg("resolved via file");
          done(JSON.parse(raw));
        } catch (e) {
          reject(e);
        }
      }
    }, 500);
    // Check immediately in case file was written before we started polling.
    if (fs.existsSync(submitFile)) {
      try {
        const raw = fs.readFileSync(submitFile, "utf-8");
        fs.unlinkSync(submitFile);
        dbg("fast-path file");
        done(JSON.parse(raw));
      } catch (e) { /* poller will retry */ }
    }
    dbg(`resolver registered, polling ${submitFile}`);
  });
}

async function handleOpenFormUi(args: { teamSiteId: string; libraryContentVersionId: string; context?: string; prefillValues?: unknown }) {
  // Push the current token to the form server so it never uses a stale value.
  try {
    await fetch(`${FORM_API_BASE}/api/set-token`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ token: currentToken }),
    });
  } catch {
    // Non-fatal — form server may not be running yet; it will fall back to its own env var.
  }

  const token = generateToken();
  const params = new URLSearchParams({
    teamSiteId: args.teamSiteId,
    versionId: args.libraryContentVersionId,
    token,
  });
  if (args.context) {
    params.set("context", Buffer.from(args.context).toString("base64"));
  }
  if (args.prefillValues) {
    params.set("prefill", Buffer.from(JSON.stringify(args.prefillValues)).toString("base64"));
  }
  const url = `${FORM_APP_BASE}/form?${params}`;
  openWithDefaultApp(url);
  return {
    url,
    token,
    message: [
      `The form has been opened in the browser: ${url}`,
      `IMPORTANT: Call get_form_result with token="${token}" after the user submits the form — NOT wait_for_form_submit (that tool is for a different code path and will hang forever here).`,
      `get_form_result polls until the form app posts the result, then returns the generatedLivedocId and outputs.`,
    ].join("\n"),
  };
}

async function handleGetFormResult(args: { token: string }) {
  const res = await fetch(`${FORM_API_BASE}/api/result/${args.token}`);
  if (res.status === 404) {
    return { error: "Result not ready yet — the form may still be open or generation is in progress. Try again in a moment." };
  }
  if (!res.ok) {
    return { error: `Failed to retrieve result (HTTP ${res.status})` };
  }
  const data = await res.json() as { generatedLivedocId: string; outputs: Array<{ id: string; status: string; format: string; fileName: string }> };
  return {
    generatedLivedocId: data.generatedLivedocId,
    outputs: data.outputs,
    hint: "Generation complete. Call get_generation_download_url with generatedLivedocId and each output's id to get download links.",
  };
}

async function handleLogin(args: {
  username?: string;
  password?: string;
  tenant?: string;
  authServiceUri?: string;
  clientId?: string;
  clientSecret?: string;
}): Promise<{ ok: boolean; message: string } | { error: string; detail: unknown }> {
  const authUri  = args.authServiceUri ?? DEFAULT_AUTH_URI;
  const tenant   = args.tenant         ?? DEFAULT_AUTH_TENANT;
  const clientId     = args.clientId     ?? DEFAULT_CLIENT_ID;
  const clientSecret = args.clientSecret ?? DEFAULT_CLIENT_SECRET;
  const username = args.username ?? DEFAULT_USERNAME;
  const password = args.password ?? DEFAULT_PASSWORD;

  if (!authUri)  return { error: "authServiceUri is required (set AUTH_SERVICE_URI env var or pass authServiceUri).", detail: null };
  if (!tenant)   return { error: "tenant is required (set AUTH_TENANT env var or pass tenant).", detail: null };
  if (!clientId) return { error: "clientId is required (set AUTH_CLIENT_ID env var or pass clientId).", detail: null };
  if (!username) return { error: "username is required (set AUTH_USERNAME env var or pass username).", detail: null };
  if (!password) return { error: "password is required (set AUTH_PASSWORD env var or pass password).", detail: null };

  const tokenUrl = `${authUri}/tenants/${encodeURIComponent(tenant)}/connect/token`;
  const body = new URLSearchParams({
    grant_type:    "client_credentials",
    client_id:     clientId,
    client_secret: clientSecret,
    username,
    password,
    scope:         LOGIN_SCOPE,
  });

  const res = await fetch(tokenUrl, {
    method:  "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body:    body.toString(),
  });

  const text = await res.text();
  let data: Record<string, unknown>;
  try { data = JSON.parse(text) as Record<string, unknown>; } catch { return { error: `Auth server returned non-JSON (HTTP ${res.status})`, detail: text }; }

  if (!res.ok || !data.access_token) {
    return { error: `Login failed (HTTP ${res.status})`, detail: data };
  }

  currentToken = data.access_token as string;
  tokenIsManual = false;
  const expiresIn = data.expires_in as number | undefined;
  return { ok: true, message: `Token obtained successfully${expiresIn ? ` (expires in ${expiresIn}s)` : ""}. All tools are now authenticated.` };
}

// ── Server wiring ───────────────────────────────────────────────────────────

const server = new McpServer({ name: "seismic-livedoc", version: "1.0.0" });

// MCP App UI resource — served when Claude Desktop opens the App panel.
// frameDomains CSP is on the registration config (resources/list) so Claude Desktop
// applies it at connection time. The shell uses callServerTool only — no connectDomains needed.
registerAppResource(
  server,
  "LiveDoc Form",
  FORM_RESOURCE_URI,
  {
    description: "LiveDoc input form — dynamically generated per template.",
    _meta: { ui: { csp: { frameDomains: ["http://127.0.0.1:3099"], connectDomains: ["http://127.0.0.1:3099"] } } },
  } as Parameters<typeof registerAppResource>[3],
  async () => {
    const bundle = await getExtAppsBundle();
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"), `[${new Date().toISOString()}] resources/read: ui://livedoc/form fetched (shell, bundle=${bundle.length}b)\n`);
    return { contents: [{
      uri: FORM_RESOURCE_URI,
      mimeType: RESOURCE_MIME_TYPE,
      text: buildShellHtml(bundle),
    }] };
  }
);

// ── Tool registrations ──────────────────────────────────────────────────────

server.registerTool(
  "search_livedoc_templates",
  {
    description:
      "Search for LiveDoc (Document Generator) templates in Seismic by name or keyword. Returns contentVersionId and teamSiteId needed for other tools. " +
      "ALWAYS call this FIRST whenever the user names or describes a template — do NOT ask for teamSiteId/libraryContentVersionId directly. " +
      "Only ask the user to disambiguate if this search returns zero or multiple plausible matches.",
    inputSchema: {
      searchText: z.string().optional().describe("Text to search across template title, description, and body."),
      page_size: z.number().optional().describe("Number of results to return (default 10, max 50)."),
    },
  },
  async (args) => {
    const result = await handleSearchTemplates(args as { searchText?: string; page_size?: number });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "search_livedoc_content",
  {
    description:
      "Search the Seismic library for content items by keyword. NOT part of the normal LiveDoc generation flow. " +
      "Only for standalone content discovery when the user explicitly asks to search for content.",
    inputSchema: {
      query: z.string().describe("Search text (title or keyword)."),
      contentType: z.enum(["ExternalSlides", "LiveSlide", "ExternalStaticSlides", "LiveDoc", "PDF"]).optional().describe("Filter by content type."),
      page_size: z.number().optional().describe("Number of results to return (default 10, max 50)."),
    },
  },
  async (args) => {
    const result = await handleSearchContent(args as { query: string; contentType?: string; page_size?: number });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

// get_livedoc_inputs — opens the MCP App panel with the form
registerAppTool(
  server,
  "get_livedoc_inputs",
  {
    description:
      "Retrieve the full input schema for a LiveDoc template and open the interactive input form in the Cowork panel. " +
      "After calling this tool the form will appear in the MCP App panel. " +
      "Immediately call wait_for_form_submit with the returned token and wait for the user to submit the form. " +
      "Do NOT build your own form, do NOT use AskUserQuestion.",
    inputSchema: {
      teamSiteId: z.string().describe("Team site identifier (UUID) that owns the template."),
      libraryContentVersionId: z.string().describe("Content version identifier (UUID) of the LiveDoc template."),
    },
    _meta: { ui: { resourceUri: FORM_RESOURCE_URI } },
  },
  async (args) => {
    const _dbgLog = (msg: string) => {
      try { fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"), `[${new Date().toISOString()}] ${msg}\n`); } catch { /* ignore */ }
    };
    _dbgLog("get_livedoc_inputs: start");
    const ir = await handleGetInputs(args as { teamSiteId: string; libraryContentVersionId: string });
    _dbgLog(`get_livedoc_inputs: handleGetInputs done, hasError=${"error" in ir}`);
    if ("error" in ir) {
      return { content: [{ type: "text" as const, text: JSON.stringify(ir, null, 2) }], isError: true };
    }

    const manualItems = (ir.manualSelectContentInput as Record<string, unknown> | undefined)
      ?.manualSelectContentItems as unknown[] | undefined ?? [];
    const isComplexForm = ir.hasImageUpload || manualItems.length > 12;
    _dbgLog(`get_livedoc_inputs: isComplexForm=${isComplexForm}`);

    if (isComplexForm) {
      return {
        content: [{ type: "text" as const, text: `This template has ${ir.hasImageUpload ? "image uploads" : `${manualItems.length} slide groups`} and requires the full Form Web App. Call open_form_ui with teamSiteId="${ir.teamSiteId}" and libraryContentVersionId="${ir.libraryContentVersionId}". After the user submits, call get_form_result (NOT wait_for_form_submit) with the token.` }],
      };
    }

    const formToken = generateToken();
    _dbgLog("get_livedoc_inputs: calling buildFormHtml");
    const formHtml = buildFormHtml(
      ir.templateName,
      ir.adhocInputs,
      ir.variableListData,
      ir.manualSelectContentInput,
      ir.forms,
      ir.teamSiteId,
      ir.libraryContentVersionId,
      formToken
    );
    _dbgLog(`get_livedoc_inputs: buildFormHtml done, htmlLen=${formHtml.length}`);

    // Register the form HTML with the local HTTP server so the App panel can iframe it.
    _dbgLog(`get_livedoc_inputs: pendingFormHtml.set ${formToken} mapSizeBefore=${pendingFormHtml.size}`);
    pendingFormHtml.set(formToken, formHtml);
    _dbgLog(`get_livedoc_inputs: pendingFormHtml.set done mapSizeAfter=${pendingFormHtml.size}`);

    // Also save to Working folder as fallback if the panel doesn't appear.
    const safeName = ir.templateName.replace(/[^a-zA-Z0-9 _-]/g, "").trim() || "livedoc-form";
    const formFileName = `${safeName} (${formToken}).html`;
    const formDir = findCoworkPath() || getDownloadsDir();
    try { fs.mkdirSync(formDir, { recursive: true }); } catch { /* best-effort */ }
    const formFilePath = uniqueFilePath(formDir, formFileName);
    _dbgLog(`get_livedoc_inputs: writing html to ${formFilePath}`);
    fs.writeFileSync(formFilePath, formHtml);
    _dbgLog("get_livedoc_inputs: returning resource response");

    const formUrl = `http://127.0.0.1:${FORM_PORT}/form/${formToken}`;
    latestFormUrl = formUrl;
    _dbgLog(`get_livedoc_inputs: latestFormUrl set to ${latestFormUrl}`);
    return {
      content: [{
        type: "text" as const,
        text: `Form loaded (token="${formToken}"). The form is now showing in the MCP App panel. Call wait_for_form_submit NOW with token="${formToken}" — do not wait for the user to say anything first; the tool will block until they submit. If the panel does not appear, the user can open **${formFileName}** from the Working folder panel instead.`,
      }],
      structuredContent: { formToken, formUrl, formHtml },
    };
  }
);

// App-only tool: returns the full form HTML for the given token so the shell can
// inject it via srcdoc (avoids HTTP requests from the sandboxed iframe).
server.registerTool(
  "get_form_html",
  {
    description: "Internal: called by the MCP App panel to retrieve the form HTML. Do NOT call this yourself.",
    inputSchema: { token: z.string() },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    const { token } = args as { token: string };
    const html = pendingFormHtml.get(token) ?? null;
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"), `[${new Date().toISOString()}] get_form_html: token=${token} found=${html !== null} len=${html?.length ?? 0} mapSize=${pendingFormHtml.size} keys=${[...pendingFormHtml.keys()].slice(0,3).join(',')}\n`);
    return {
      content: [{ type: "text" as const, text: html ?? "" }],
      structuredContent: { html, token },
    };
  }
);

server.registerTool(
  "log_debug_message",
  {
    description: "Internal: debug logging from the App panel. Do NOT call this yourself.",
    inputSchema: { msg: z.string() },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"),
      `[${new Date().toISOString()}] APP_PANEL: ${(args as { msg: string }).msg}\n`);
    return { content: [{ type: "text" as const, text: "ok" }] };
  }
);

server.registerTool(
  "receive_form_submission",
  {
    description: "Internal: relays the App panel form submission to the form server. Do NOT call this yourself.",
    inputSchema: {
      token: z.string(),
      payload: z.string().describe("JSON-serialized form payload"),
    },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    const { token, payload } = args as { token: string; payload: string };
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"),
      `[${new Date().toISOString()}] receive_form_submission: token=${token} payloadLen=${payload.length}\n`);
    // First try in-process resolution (if running in the same process as wait_for_form_submit).
    const resolver = pendingForms.get(token);
    if (resolver) {
      pendingForms.delete(token);
      pendingFormHtml.delete(token);
      try { resolver(JSON.parse(payload)); } catch { resolver(payload); }
      fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"),
        `[${new Date().toISOString()}] receive_form_submission: resolved in-process\n`);
      return { content: [{ type: "text" as const, text: "ok" }] };
    }
    // Cross-process fallback: write a temp file that wait_for_form_submit polls.
    // Two separate Node.js processes are spawned by Claude Desktop — they share no memory,
    // and only one of them wins port 3099, so HTTP relay is unreliable.
    const submitFile = path.join(os.tmpdir(), `livedoc-submit-${token}.json`);
    fs.writeFileSync(submitFile, payload, "utf-8");
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"),
      `[${new Date().toISOString()}] receive_form_submission: wrote file ${submitFile}\n`);
    return { content: [{ type: "text" as const, text: "ok" }] };
  }
);

server.registerTool(
  "wait_for_form_submit",
  {
    description:
      "Wait (up to 10 minutes) for the user to fill and submit the LiveDoc input form. Returns the form payload — pass it directly to submit_livedoc_generation. Call this immediately after get_livedoc_inputs.",
    inputSchema: {
      token: z.string().describe("The form session token returned by get_livedoc_inputs."),
    },
  },
  async (args) => {
    const payload = await handleWaitForFormSubmit(args as { token: string });
    return {
      content: [{
        type: "text" as const,
        text: [
          "Form submitted. NEXT: call submit_livedoc_generation immediately with this exact payload — do not modify it:",
          JSON.stringify(payload, null, 2),
        ].join("\n"),
      }],
    };
  }
);

server.registerTool(
  "submit_livedoc_generation",
  {
    description:
      "Submit a LiveDoc generation job. Provide ad hoc input values and at least one output format (PPTX, DOCX, PDF). Returns a generatedLivedocId to poll for status.",
    inputSchema: {
      teamSiteId: z.string().describe("Team site identifier (UUID)."),
      libraryContentVersionId: z.string().describe("Content version identifier (UUID) of the LiveDoc template."),
      adHocInputs: z.array(z.object({ name: z.string(), value: z.any() })).describe("Array of {name, value} pairs for ALL ad hoc inputs."),
      outputs: z.array(z.object({
        format: z.string(),
        name: z.string().optional(),
        fileName: z.string().describe("Filename with extension, e.g. \"Template.pdf\". Always set this."),
      })).describe("Output formats to generate."),
      variableListData: z.array(z.object({
        variableListName: z.string(),
        variableInputs: z.array(z.object({ name: z.string(), value: z.any() })),
      })).optional().describe("Variable list data from variableListDefinitions."),
      liveFormSellerTemplateId: z.string().optional(),
      regionalFormat: z.string().optional().describe("Regional format culture name, e.g. \"en-US\"."),
      manualSelectContentInput: z.object({
        manualSelectContentItems: z.array(z.object({
          id: z.string().describe("Stable slot identifier — copy verbatim from get_livedoc_inputs."),
          name: z.string().optional(),
          contentType: z.string().describe("One of \"Group\", \"Section\", \"LiveSlide\", \"ResourcePDF\", etc."),
          versionId: z.string().optional().describe("contentVersionId from wait_for_form_submit. Never populate yourself."),
          sourceBlobId: z.string().optional(),
          pageNumber: z.number().optional(),
          isInclude: z.boolean(),
          orderIndex: z.number().optional(),
        })),
      }).optional().describe("Content selection — pass ONLY what wait_for_form_submit returned."),
    },
  },
  async (args) => {
    const subResult = await handleSubmitGeneration(args as Parameters<typeof handleSubmitGeneration>[0]);
    const subBody = subResult as Record<string, unknown>;
    if (subBody.error) {
      return { content: [{ type: "text" as const, text: JSON.stringify(subBody, null, 2) }], isError: true };
    }
    const gid = String(subBody.generatedLivedocId ?? "");
    return {
      content: [{
        type: "text" as const,
        text: [
          `Generation started. generatedLivedocId: ${gid}`,
          `NEXT: call get_generation_status with generatedLivedocId="${gid}". Keep calling every few seconds until allDone=true.`,
        ].join("\n"),
      }],
    };
  }
);

server.registerTool(
  "get_generation_status",
  {
    description: "Check the status of a LiveDoc generation job. Poll until all outputs reach 'Completed' or 'Failed'.",
    inputSchema: {
      generatedLivedocId: z.string().describe("The generatedLivedocId returned by submit_livedoc_generation."),
    },
  },
  async (args) => {
    const statusResult = await handleGetStatus(args as { generatedLivedocId: string });
    const st = statusResult as Record<string, unknown>;
    if (st.error) {
      return { content: [{ type: "text" as const, text: JSON.stringify(st, null, 2) }], isError: true };
    }
    const outputs = (st.outputs as Array<Record<string, unknown>>) ?? [];
    const allDone = st.allDone as boolean;
    const nextStep = allDone
      ? `All done. NEXT: call download_generation_output for each completed output:\n${outputs.filter(o => o.status === "Completed").map(o => `  outputId="${o.id}" (${o.format} — ${o.fileName})`).join("\n")}`
      : `Still generating. NEXT: call get_generation_status again with generatedLivedocId="${st.generatedLivedocId}" in a few seconds.`;
    return {
      content: [{
        type: "text" as const,
        text: [JSON.stringify(st, null, 2), nextStep].join("\n\n"),
      }],
    };
  }
);

server.registerTool(
  "open_form_ui",
  {
    description:
      "Opens the LiveDoc Form Web App for templates that require image uploads. Call this ONLY when get_livedoc_inputs explicitly instructs you to (hasImageUpload case). After submission, call get_form_result (NOT wait_for_form_submit) with the returned token.",
    inputSchema: {
      teamSiteId: z.string().describe("Team site identifier (UUID)."),
      libraryContentVersionId: z.string().describe("Content version identifier (UUID) of the LiveDoc template."),
      context: z.string().optional().describe("The user's original generation request (natural language)."),
      prefillValues: z.any().optional().describe("Optional AI-suggested default values to pre-populate the form."),
    },
  },
  async (args) => {
    const result = await handleOpenFormUi(args as { teamSiteId: string; libraryContentVersionId: string; context?: string; prefillValues?: unknown });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "get_form_result",
  {
    description: "Retrieve the generation result posted back by the Form UI after the user completed and closed the form.",
    inputSchema: {
      token: z.string().describe("The token returned by open_form_ui."),
    },
  },
  async (args) => {
    const result = await handleGetFormResult(args as { token: string });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "login",
  {
    description:
      "Obtain a Seismic bearer token using username and password (OAuth 2.0 Resource Owner Password Credentials). Automatically sets the token for all subsequent tool calls.",
    inputSchema: {
      username: z.string().optional().describe(`Seismic username. Defaults to AUTH_USERNAME env var${DEFAULT_USERNAME ? " (pre-configured)" : ""}.`),
      password: z.string().optional().describe(`Seismic password. Defaults to AUTH_PASSWORD env var.`),
      tenant: z.string().optional().describe(`Tenant slug, e.g. "qa01eastasia01". Defaults to AUTH_TENANT env var (${DEFAULT_AUTH_TENANT || "not set"}).`),
      authServiceUri: z.string().optional().describe("Auth service base URL. Defaults to AUTH_SERVICE_URI env var."),
      clientId: z.string().optional().describe("OAuth client ID."),
      clientSecret: z.string().optional().describe("OAuth client secret."),
    },
  },
  async (args) => {
    const result = await handleLogin(args as Parameters<typeof handleLogin>[0]);
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "set_token",
  {
    description: "Update the Seismic API bearer token used by all other tools. Call this whenever a tool returns HTTP 401 before retrying.",
    inputSchema: {
      token: z.string().describe("The new bearer token (without the 'Bearer ' prefix)."),
    },
  },
  async (args) => {
    currentToken = (args as { token: string }).token;
    tokenIsManual = true;
    return { content: [{ type: "text" as const, text: "Token updated. This token will not be auto-replaced by the credential-flow login on a 401." }] };
  }
);

server.registerTool(
  "get_generation_download_url",
  {
    description:
      "Get the download URL for a completed LiveDoc output. Returns a JSON payload with the URL. Use outputId from get_generation_status, or a format alias like 'pptx', 'docx', 'pdf'.",
    inputSchema: {
      generatedLivedocId: z.string().describe("The generatedLivedocId."),
      outputId: z.string().describe("The output id or a format alias: 'pptx', 'docx', 'pdf', 'gslides', 'gdoc'."),
    },
  },
  async (args) => {
    const dlResult = await handleGetDownloadUrl(args as { generatedLivedocId: string; outputId: string });
    if (dlResult && typeof dlResult === "object" && "error" in (dlResult as object)) {
      return { content: [{ type: "text" as const, text: JSON.stringify(dlResult, null, 2) }], isError: true };
    }
    const dlBody = dlResult as Record<string, unknown>;
    const dlUrl = String(dlBody.url ?? dlBody.downloadUrl ?? dlBody.Url ?? dlBody.DownloadUrl ?? "");
    const dlFile = String(dlBody.fileName ?? dlBody.FileName ?? dlBody.name ?? dlBody.Name ?? "download");
    if (!dlUrl) {
      return { content: [{ type: "text" as const, text: JSON.stringify(dlResult, null, 2) }] };
    }
    return {
      content: [{
        type: "text" as const,
        text: `✅ **${dlFile}** is ready.\n\nDownload link: [${dlFile}](${dlUrl})\n\n(Reproduce the markdown link above verbatim in your reply so the user can click it.)`,
      }],
    };
  }
);

server.registerTool(
  "download_generation_output",
  {
    description:
      "Download a completed LiveDoc output to the Working folder and open it with the system default app.",
    inputSchema: {
      generatedLivedocId: z.string().describe("The generatedLivedocId."),
      outputId: z.string().describe("The output id or a format alias: 'pptx', 'docx', 'pdf', 'gslides', 'gdoc'."),
      autoOpen: z.boolean().optional().describe("Whether to automatically open the file. Default true."),
    },
  },
  async (args) => {
    const result = await handleDownloadGenerationOutput(args as { generatedLivedocId: string; outputId: string; autoOpen?: boolean });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "debug_environment",
  {
    description:
      "Diagnostic tool: dumps process working directory, matched environment variable names/values, and client capabilities.",
    inputSchema: {},
  },
  async () => {
    const keywords = /claude|anthropic|output|sandbox|session|cowork|agent|workspace/i;
    const looksLikePath = (v: string) => /^[a-zA-Z]:[\\/]|^\//.test(v) && /[\\/]/.test(v);
    const matchedVarNames = Object.keys(process.env).filter((k) => keywords.test(k));
    const pathLikeValues = Object.fromEntries(
      matchedVarNames
        .map((k) => [k, process.env[k] ?? ""])
        .filter(([, v]) => looksLikePath(v as string))
    );
    const clientCaps = server.server.getClientCapabilities();
    const result = {
      cwd: process.cwd(),
      matchedEnvVarNames: matchedVarNames,
      pathLikeEnvVarValues: pathLikeValues,
      clientCapabilitiesRawJSON: JSON.stringify(clientCaps, null, 2),
      elicitationSupported: !!(clientCaps?.elicitation),
      mcpAppUiSupported: !!(clientCaps && JSON.stringify(clientCaps).includes("mcp-app")),
    };
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

// ── PPTX Auto-Tagging tools (PoC) ────────────────────────────────────────────

server.registerTool(
  "pptx_extract_shapes",
  {
    description:
      "Extract all shapes from a PPTX file and return a structured list. Requires the local PoC server (cd poc-auto-tagging && npm start).",
    inputSchema: {
      pptxBase64: z.string().describe("Base64-encoded PPTX file content."),
      slideIndex: z.number().optional().describe("Optional: only extract shapes from this slide (0-based)."),
    },
  },
  async (args) => {
    const a = args as { pptxBase64: string; slideIndex?: number };
    const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
    const r = await fetch(`${pocBase}/api/pptx/extract`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pptxBase64: a.pptxBase64, ...(a.slideIndex !== undefined ? { slideIndex: a.slideIndex } : {}) }),
    });
    if (!r.ok) throw new Error(`pptx_extract_shapes HTTP ${r.status}: ${await r.text()}`);
    return { content: [{ type: "text" as const, text: JSON.stringify(await r.json(), null, 2) }] };
  }
);

server.registerTool(
  "pptx_auto_tag",
  {
    description:
      "Send a PPTX to the AI (local LLM) for automatic analysis. Requires the local PoC server (cd poc-auto-tagging && npm start).",
    inputSchema: {
      pptxBase64: z.string().describe("Base64-encoded PPTX file content."),
      schema: z.record(z.string(), z.string()).optional().describe("Optional datasource schema as a JSON object."),
    },
  },
  async (args) => {
    const a = args as { pptxBase64: string; schema?: Record<string, string> };
    const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
    const r = await fetch(`${pocBase}/api/pptx/auto-tag`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pptxBase64: a.pptxBase64, schema: a.schema ?? {} }),
    });
    if (!r.ok) throw new Error(`pptx_auto_tag HTTP ${r.status}: ${await r.text()}`);
    return { content: [{ type: "text" as const, text: JSON.stringify(await r.json(), null, 2) }] };
  }
);

server.registerTool(
  "pptx_mark_shapes",
  {
    description:
      "Apply dynamic element markings to a PPTX file. Requires the local PoC server (cd poc-auto-tagging && npm start).",
    inputSchema: {
      pptxBase64: z.string().describe("Base64-encoded PPTX file content."),
      marks: z.array(z.object({
        slideIndex: z.number().describe("0-based slide index."),
        shapeId: z.number().describe("Numeric shape ID from pptx_extract_shapes."),
        varName: z.string().describe("camelCase variable name, e.g. 'companyName'."),
        varType: z.enum(["text", "image", "table", "chart", "number", "date"]),
        description: z.string().optional(),
      })).describe("List of shapes to mark as dynamic."),
    },
  },
  async (args) => {
    const a = args as { pptxBase64: string; marks: unknown[] };
    const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
    const r = await fetch(`${pocBase}/api/pptx/mark`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pptxBase64: a.pptxBase64, marks: a.marks }),
    });
    if (!r.ok) throw new Error(`pptx_mark_shapes HTTP ${r.status}: ${await r.text()}`);
    const result = await r.json() as { bindings: unknown[]; markedAt: string };
    return {
      content: [{
        type: "text" as const,
        text: JSON.stringify({ markedAt: result.markedAt, bindingCount: result.bindings.length, bindings: result.bindings }, null, 2) +
          "\n\n(markedPptxBase64 available in full result — use pptx_mark_shapes result.pptxBase64 to save the file)",
      }],
    };
  }
);

server.registerTool(
  "pptx_get_manifest",
  {
    description:
      "Extract the current binding manifest from a PPTX file. Requires the local PoC server (cd poc-auto-tagging && npm start).",
    inputSchema: {
      pptxBase64: z.string().describe("Base64-encoded PPTX file content."),
    },
  },
  async (args) => {
    const a = args as { pptxBase64: string };
    const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
    const r = await fetch(`${pocBase}/api/pptx/manifest?pptxBase64=${encodeURIComponent(a.pptxBase64)}`);
    if (!r.ok) throw new Error(`pptx_get_manifest HTTP ${r.status}: ${await r.text()}`);
    return { content: [{ type: "text" as const, text: JSON.stringify(await r.json(), null, 2) }] };
  }
);

process.on("SIGTERM", () => process.exit(0));
process.on("SIGINT", () => process.exit(0));
process.on("uncaughtException", (err) => {
  fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-crash.log"), `[${new Date().toISOString()}] uncaughtException: ${err.stack ?? err.message}\n`);
  process.exit(1);
});
process.on("unhandledRejection", (reason) => {
  fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-crash.log"), `[${new Date().toISOString()}] unhandledRejection: ${reason}\n`);
  process.exit(1);
});

// Auto-login on startup when credentials are available via env vars
if (!currentToken && DEFAULT_USERNAME && DEFAULT_PASSWORD) {
  await autoLogin();
}

const debugLog = path.join(os.tmpdir(), "mcp-livedoc-debug.log");
server.server.oninitialized = () => {
  const caps = server.server.getClientCapabilities();
  const uiCap = getUiCapability(caps as Parameters<typeof getUiCapability>[0]);
  fs.appendFileSync(debugLog, `[${new Date().toISOString()}] oninitialized: extensions=${JSON.stringify(caps?.extensions)}, uiCap=${JSON.stringify(uiCap)}\n`);
};

const transport = new StdioServerTransport();
await server.connect(transport);
