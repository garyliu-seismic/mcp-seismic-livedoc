import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { validatePrefillFieldNames } from "./prefillValidation.js";

const schema = {
  adhocScalars: [{ name: "CompanyName" }, { name: "Amount" }],
  adhocTables: [{ name: "LineItems" }],
  variableLists: [{ name: "Signers" }],
};

describe("validatePrefillFieldNames", () => {
  test("accepts field names that exist in the schema", () => {
    const result = validatePrefillFieldNames(schema, {
      scalars: { CompanyName: "Acme" },
      tables: { LineItems: [{ Qty: 1 }] },
      variableLists: { Signers: { scalars: { Name: "Jo" } } },
    });
    assert.deepEqual(result.unknown, []);
  });

  test("flags an unknown scalar name", () => {
    const result = validatePrefillFieldNames(schema, { scalars: { Bogus: "x" } });
    assert.deepEqual(result.unknown, ["scalars.Bogus"]);
  });

  test("flags unknown table and variable-list names together", () => {
    const result = validatePrefillFieldNames(schema, {
      tables: { NotARealTable: [] },
      variableLists: { NotARealVL: {} },
    });
    assert.deepEqual(result.unknown.sort(), ["tables.NotARealTable", "variableLists.NotARealVL"]);
  });

  test("reports the valid name lists back so the caller can echo them in an error", () => {
    const result = validatePrefillFieldNames(schema, {});
    assert.deepEqual(result.validScalars, ["CompanyName", "Amount"]);
    assert.deepEqual(result.validTables, ["LineItems"]);
    assert.deepEqual(result.validVariableLists, ["Signers"]);
  });

  test("treats a schema with no fields as valid but empty", () => {
    const result = validatePrefillFieldNames({}, { scalars: { Anything: 1 } });
    assert.deepEqual(result.unknown, ["scalars.Anything"]);
    assert.deepEqual(result.validScalars, []);
  });
});
