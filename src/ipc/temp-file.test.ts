import { test, describe } from "node:test";
import assert from "node:assert/strict";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import {
  writeSchema, readSchema, deleteSchemaFile,
  writeLatestToken, readLatestToken,
  writeResult, readResult, deleteResultFile,
  writePrefill, readPrefill, deletePrefillFile,
  writeGid, readGid, deleteGidFile,
} from "./temp-file.js";

// Use a per-test-run token prefix so parallel/re-run test invocations never collide
// on the same real %TEMP% files.
const runId = `test-${process.pid}-${Date.now()}`;

describe("schema file IPC", () => {
  test("writeSchema/readSchema round-trips through disk", () => {
    const token = `${runId}-schema`;
    const schema = { adhocScalars: [{ name: "Foo", type: "string" }] };
    writeSchema(token, schema);
    assert.deepEqual(readSchema(token), schema);
  });

  test("readSchema returns null for an unknown token", () => {
    assert.equal(readSchema(`${runId}-never-written`), null);
  });

  test("deleteSchemaFile removes the file from disk (but not the in-memory cache)", () => {
    const token = `${runId}-schema-delete`;
    writeSchema(token, { ok: true });
    const filePath = path.join(os.tmpdir(), `mcp-livedoc-schema-${token}.json`);
    assert.equal(fs.existsSync(filePath), true);
    deleteSchemaFile(token);
    assert.equal(fs.existsSync(filePath), false);
  });
});

describe("latest-token pointer", () => {
  test("writeLatestToken/readLatestToken round-trips", () => {
    const token = `${runId}-latest`;
    writeLatestToken(token);
    const data = readLatestToken();
    assert.equal(data?.formToken, token);
    assert.equal(typeof data?.writtenAt, "number");
  });
});

describe("result file IPC", () => {
  test("writeResult/readResult round-trips", () => {
    const token = `${runId}-result`;
    const data = { generatedLivedocId: "gid-1", status: "Completed", downloadUrls: ["http://x"] };
    writeResult(token, data);
    assert.deepEqual(readResult(token), data);
  });

  test("readResult returns null once deleteResultFile has run", () => {
    const token = `${runId}-result-delete`;
    writeResult(token, { generatedLivedocId: "gid-2", status: "Completed", downloadUrls: [] });
    deleteResultFile(token);
    assert.equal(readResult(token), null);
  });
});

describe("prefill file IPC", () => {
  test("writePrefill/readPrefill round-trips", () => {
    const token = `${runId}-prefill`;
    const prefill = { scalars: { Name: "Acme" } };
    writePrefill(token, prefill);
    assert.deepEqual(readPrefill(token), prefill);
  });

  test("deletePrefillFile removes the file", () => {
    const token = `${runId}-prefill-delete`;
    writePrefill(token, { scalars: { A: 1 } });
    deletePrefillFile(token);
    assert.equal(readPrefill(token), null);
  });
});

describe("gid → formToken mapping", () => {
  test("writeGid/readGid round-trips", () => {
    const gid = `${runId}-gid`;
    writeGid(gid, "form-token-123");
    assert.deepEqual(readGid(gid), { formToken: "form-token-123" });
  });

  test("deleteGidFile removes the mapping so a later poll finds nothing", () => {
    const gid = `${runId}-gid-delete`;
    writeGid(gid, "form-token-456");
    deleteGidFile(gid);
    assert.equal(readGid(gid), null);
  });
});
