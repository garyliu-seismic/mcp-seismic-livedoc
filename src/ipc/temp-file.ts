import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { dbg } from "../utils/debug.js";
import type { LatestTokenData, ResultFileData, PrefillData } from "../types.js";

export const pendingFormSchemas = new Map<string, unknown>(); // token â†' normalised form schema for get_form_schema
export const pendingGenerations = new Map<string, string>();  // generatedLivedocId â†' formToken (Process 2 only)

const schemaPath     = (token: string) => path.join(os.tmpdir(), `mcp-livedoc-schema-${token}.json`);
const latestTokenPath = ()             => path.join(os.tmpdir(), "mcp-livedoc-latest-token.json");
const resultPath     = (token: string) => path.join(os.tmpdir(), `mcp-livedoc-result-${token}.json`);
const gidPath        = (gid: string)   => path.join(os.tmpdir(), `mcp-livedoc-gid-${gid}.json`);
const prefillPath    = (token: string) => path.join(os.tmpdir(), `mcp-livedoc-prefill-${token}.json`);

export function writeSchema(token: string, schema: unknown): void {
  pendingFormSchemas.set(token, schema);
  try {
    fs.writeFileSync(schemaPath(token), JSON.stringify(schema), "utf-8");
  } catch (e) { dbg(`writeSchema error: ${e}`); }
}

export function readSchema(token: string): unknown | null {
  const cached = pendingFormSchemas.get(token);
  if (cached !== undefined) return cached;
  try {
    const p = schemaPath(token);
    if (!fs.existsSync(p)) return null;
    const schema = JSON.parse(fs.readFileSync(p, "utf-8")) as unknown;
    pendingFormSchemas.set(token, schema);
    return schema;
  } catch { return null; }
}

export function deleteSchemaFile(token: string): void {
  try { fs.unlinkSync(schemaPath(token)); } catch { /* ignore */ }
}

export function writeLatestToken(formToken: string): void {
  try {
    fs.writeFileSync(latestTokenPath(), JSON.stringify({ formToken, writtenAt: Date.now() }), "utf-8");
  } catch (e) { dbg(`writeLatestToken error: ${e}`); }
}

export function readLatestToken(): (LatestTokenData & { writtenAt?: number }) | null {
  try {
    const p = latestTokenPath();
    if (!fs.existsSync(p)) return null;
    return JSON.parse(fs.readFileSync(p, "utf-8")) as LatestTokenData & { writtenAt?: number };
  } catch { return null; }
}

export function writeResult(token: string, data: ResultFileData): void {
  try {
    fs.writeFileSync(resultPath(token), JSON.stringify(data), "utf-8");
  } catch (e) { dbg(`writeResult error: ${e}`); }
}

export function readResult(token: string): ResultFileData | null {
  try {
    const p = resultPath(token);
    if (!fs.existsSync(p)) return null;
    return JSON.parse(fs.readFileSync(p, "utf-8")) as ResultFileData;
  } catch { return null; }
}

export function writePrefill(token: string, prefill: PrefillData): void {
  try {
    fs.writeFileSync(prefillPath(token), JSON.stringify(prefill), "utf-8");
  } catch (e) { dbg(`writePrefill error: ${e}`); }
}

export function readPrefill(token: string): PrefillData | null {
  try {
    const p = prefillPath(token);
    if (!fs.existsSync(p)) return null;
    return JSON.parse(fs.readFileSync(p, "utf-8")) as PrefillData;
  } catch { return null; }
}

export function deletePrefillFile(token: string): void {
  try { fs.unlinkSync(prefillPath(token)); } catch { /* ignore */ }
}

export function writeGid(gid: string, formToken: string): void {
  try {
    fs.writeFileSync(gidPath(gid), JSON.stringify({ formToken }), "utf-8");
  } catch (e) { dbg(`writeGid error: ${e}`); }
}

export function readGid(gid: string): { formToken: string } | null {
  try {
    const p = gidPath(gid);
    if (!fs.existsSync(p)) return null;
    return JSON.parse(fs.readFileSync(p, "utf-8")) as { formToken: string };
  } catch { return null; }
}
