#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { getUiCapability } from "@modelcontextprotocol/ext-apps/server";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { dbg } from "./utils/debug.js";
import { loadSavedToken } from "./auth/token-store.js";
import { setToken, getToken, isTokenManual } from "./auth/state.js";
import { autoLogin } from "./auth/auto-login.js";
import { jwtExpiresAt } from "./auth/jwt.js";
import { registerChatTools } from "./tools/chat-tools.js";
import { registerPanelTools } from "./tools/panel-tools.js";
import { registerDeprecatedTools } from "./tools/deprecated-tools.js";
import { registerPptxAutotagTools } from "./tools/pptx-autotag.js";

const server = new McpServer({ name: "seismic-livedoc", version: "1.0.0" });

registerChatTools(server);
registerPanelTools(server);
registerDeprecatedTools(server);
registerPptxAutotagTools(server);

// Process signal handlers
process.on("SIGTERM", () => process.exit(0));
process.on("SIGINT",  () => process.exit(0));
process.on("uncaughtException", (err) => {
  fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-crash.log"), `[${new Date().toISOString()}] uncaughtException: ${err.stack ?? err.message}\n`);
  process.exit(1);
});
process.on("unhandledRejection", (reason) => {
  fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-crash.log"), `[${new Date().toISOString()}] unhandledRejection: ${reason}\n`);
  process.exit(1);
});

// Token initialisation on startup (preserve exact existing logic)
const envToken = process.env.SEISMIC_API_TOKEN ?? "";
const savedToken = loadSavedToken();
if (savedToken && (!envToken || savedToken !== envToken)) {
  setToken(savedToken);
  dbg(`startup: using saved token (expires ${(() => { const e = jwtExpiresAt(savedToken); return e ? new Date(e).toISOString() : "unknown"; })()})`);
} else if (envToken) {
  setToken(envToken);
}
if (!getToken() || (jwtExpiresAt(getToken()) !== null && Date.now() >= (jwtExpiresAt(getToken()) ?? 0) - 60_000)) {
  autoLogin().then(ok => { if (ok) dbg("startup: autoLogin succeeded"); }).catch(() => {});
}

// Proactively refresh the token when it is within 5 minutes of expiry.
// Runs every 60 s; skips when token was set manually (set_token tool) to
// avoid clobbering an explicitly-provided token.
setInterval(() => {
  if (isTokenManual()) return;
  const exp = jwtExpiresAt(getToken());
  if (exp === null) return; // non-JWT token — don't touch it
  const msRemaining = exp - Date.now();
  if (msRemaining > 0 && msRemaining <= 5 * 60_000) {
    dbg(`token-refresh: expires in ${Math.round(msRemaining / 1000)}s, refreshing`);
    autoLogin().then(ok => { if (ok) dbg("token-refresh: succeeded"); }).catch(() => {});
  }
}, 60_000).unref?.();

const transport = new StdioServerTransport();
server.server.oninitialized = () => {
  const uiCap = getUiCapability(server.server);
  const caps = (server.server as unknown as { _clientCapabilities?: unknown })._clientCapabilities;
  dbg(`oninitialized: extensions=${JSON.stringify(caps && typeof caps === 'object' && 'extensions' in caps ? (caps as Record<string, unknown>).extensions : null)}, uiCap=${JSON.stringify(uiCap)}`);
};
await server.connect(transport);
