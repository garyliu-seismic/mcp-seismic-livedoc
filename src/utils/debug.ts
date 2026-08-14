import * as fs from "fs";
import * as path from "path";
import * as os from "os";

export const DEBUG_LOG = path.join(os.tmpdir(), "mcp-livedoc-debug.log");
export const DEBUG_ENABLED = !!process.env.MCP_LIVEDOC_DEBUG;

export function dbg(msg: string): void {
  if (!DEBUG_ENABLED) return;
  try { fs.appendFileSync(DEBUG_LOG, `[${new Date().toISOString()}] ${msg}\n`); } catch {}
}
