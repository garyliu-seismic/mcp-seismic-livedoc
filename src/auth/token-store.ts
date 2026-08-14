import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { dbg } from "../utils/debug.js";
import { jwtExpiresAt } from "./jwt.js";

// â"€â"€ Token persistence â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
// Panel-login tokens are stored here so they survive MCP server restarts.
export const TOKEN_FILE = path.join(os.tmpdir(), "mcp-livedoc-token.json");

export function saveToken(token: string): void {
  try {
    fs.writeFileSync(TOKEN_FILE, JSON.stringify({ token }), { encoding: "utf-8", mode: 0o600 });
  } catch (e) { dbg(`saveToken error: ${e}`); }
}

export function loadSavedToken(): string {
  try {
    const { token } = JSON.parse(fs.readFileSync(TOKEN_FILE, "utf-8")) as { token: string };
    if (!token) return "";
    const exp = jwtExpiresAt(token);
    if (exp !== null && Date.now() >= exp - 60_000) { dbg("loadSavedToken: found token but it is expired, skipping"); return ""; }
    dbg(`loadSavedToken: loaded valid token (expires ${exp ? new Date(exp).toISOString() : "never"})`);
    return token;
  } catch { return ""; }
}
