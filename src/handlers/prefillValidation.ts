import type { PrefillData } from "../types.js";

export interface PrefillableSchema {
  adhocScalars?: Array<{ name: string }>;
  adhocTables?: Array<{ name: string }>;
  variableLists?: Array<{ name: string }>;
}

export interface PrefillValidationResult {
  unknown: string[];
  validScalars: string[];
  validTables: string[];
  validVariableLists: string[];
}

/**
 * Checks the field/table/variable-list names in a prefill request against the
 * names actually present in the form schema. Kept pure (no I/O) so it can be
 * unit-tested without spinning up the MCP server or writing temp files.
 */
export function validatePrefillFieldNames(
  schema: PrefillableSchema,
  request: Pick<PrefillData, "scalars" | "tables" | "variableLists">
): PrefillValidationResult {
  const validScalars = new Set((schema.adhocScalars ?? []).map((f) => f.name));
  const validTables = new Set((schema.adhocTables ?? []).map((t) => t.name));
  const validVLs = new Set((schema.variableLists ?? []).map((vl) => vl.name));

  const unknown: string[] = [];
  Object.keys(request.scalars ?? {}).forEach((k) => { if (!validScalars.has(k)) unknown.push(`scalars.${k}`); });
  Object.keys(request.tables ?? {}).forEach((k) => { if (!validTables.has(k)) unknown.push(`tables.${k}`); });
  Object.keys(request.variableLists ?? {}).forEach((k) => { if (!validVLs.has(k)) unknown.push(`variableLists.${k}`); });

  return {
    unknown,
    validScalars: [...validScalars],
    validTables: [...validTables],
    validVariableLists: [...validVLs],
  };
}
