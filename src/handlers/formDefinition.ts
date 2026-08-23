import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { apiFetch, gf } from "../api/client.js";
import { DEBUG_LOG } from "../utils/debug.js";

// Unconditional (not gated behind MCP_LIVEDOC_DEBUG) so success/failure is always visible
// in the log — matches the existing POLL: convention in panel-tools.ts.
function log(msg: string): void {
  try { fs.appendFileSync(DEBUG_LOG, `[${new Date().toISOString()}] FORMDEF: ${msg}\n`); } catch { /* best-effort */ }
}

export interface FieldValidation {
  required?: boolean;
  // Raw ValidationSetting.ValidWhen expression from the form definition — Seismic's form
  // builder only exposes required-ness as a simple flag; anything more (length/range/regex)
  // is expressed as this conditional formula rather than fielded properties. Not evaluated
  // client-side (format unconfirmed) — surfaced as a hint only.
  formula?: unknown;
}

export interface FieldMeta {
  validation?: FieldValidation;
  // Static domain-of-value: a fixed list of choices embedded directly in the form
  // definition (no server round-trip needed). Dynamic DOVs (backed by a data source /
  // requiring RetrieveFormData?type=DOV) are intentionally NOT surfaced here.
  options?: Array<{ value: string; label: string }>;
  // From BasicSetting — the template author's own display label/help text/tooltip, as an
  // alternative to showing the raw bound variable name.
  label?: string;
  helpText?: string;
  tooltip?: string;
  // From DataSetting.DefaultValue — only surfaced when DefaultValueType is "Custom" (a
  // literal), never for formula/variable-driven defaults, which aren't a value we can prefill.
  defaultValue?: unknown;
}

/**
 * GET /v3/forms/{formId} — the form definition (FormElements, validation rules,
 * domain-of-value lists), as opposed to the higher-level /v3/teamsites/.../livedocVersions/...
 * summary already used by handleGetInputs.
 */
export async function fetchFormDefinition(formId: string, teamSiteId: string) {
  log(`request formId=${formId} teamSiteId=${teamSiteId}`);
  const result = await apiFetch(`/v3/forms/${encodeURIComponent(formId)}?teamSiteId=${encodeURIComponent(teamSiteId)}`);
  if (result.status === 200) {
    log(`success formId=${formId} HTTP 200`);
  } else {
    log(`FAILED formId=${formId} HTTP ${result.status} body=${JSON.stringify(result.body).slice(0, 300)}`);
  }
  return result;
}

// A DOV is dynamic (needs RetrieveFormData?type=DOV to resolve) if it references a data
// source. Anything else with an inline value list is static.
function isDynamicDov(dov: Record<string, unknown>): boolean {
  return !!(gf(dov, "dataSourceId") || gf(dov, "dataSourceName") || gf(dov, "isDynamic"));
}

// Turns a raw value-list container (array of items, or an object wrapping one under
// Items/Values/Options) into {value,label} pairs, unless it looks data-source-backed.
function toOptions(raw: unknown): Array<{ value: string; label: string }> | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  let list: unknown;
  if (Array.isArray(raw)) {
    list = raw;
  } else {
    const o = raw as Record<string, unknown>;
    if (isDynamicDov(o)) return undefined;
    list = gf(o, "items") ?? gf(o, "values") ?? gf(o, "options") ?? gf(o, "choices");
  }
  if (!Array.isArray(list) || list.length === 0) return undefined;
  const options = list.map(v => {
    if (v && typeof v === "object") {
      const o = v as Record<string, unknown>;
      const value = String(gf(o, "value") ?? gf(o, "key") ?? gf(o, "name") ?? "");
      const label = String(gf(o, "label") ?? gf(o, "displayName") ?? gf(o, "text") ?? gf(o, "name") ?? value);
      return { value, label };
    }
    return { value: String(v), label: String(v) };
  }).filter(o => o.value !== "");
  return options.length > 0 ? options : undefined;
}

// Per the live GET /v3/forms/{formId} response (2026-08-22): each element has
// DataSetting.DomainOfValue (null when unset) and TypeSpecificSetting.{ComboBox,ChoicesList,
// RadioButton,SingleSelectionList,AutoCompleteComboBox} — any of these may carry an inline
// static option list. No sample with a populated list was available yet, so this tries every
// plausible container; verify against mcp-livedoc-debug-form-definition.json once a template
// with an actual dropdown/DOV field is inspected.
function readStaticOptions(el: Record<string, unknown>): Array<{ value: string; label: string }> | undefined {
  const dataSetting = gf(el, "dataSetting") as Record<string, unknown> | undefined;
  if (dataSetting) {
    const fromDov = toOptions(gf(dataSetting, "domainOfValue"));
    if (fromDov) return fromDov;
  }
  const typeSpecific = gf(el, "typeSpecificSetting") as Record<string, unknown> | undefined;
  if (typeSpecific) {
    for (const key of ["comboBox", "choicesList", "radioButton", "singleSelectionList", "autoCompleteComboBox"]) {
      const fromWidget = toOptions(gf(typeSpecific, key));
      if (fromWidget) return fromWidget;
    }
  }
  return undefined;
}

// Per the live response: required-ness is ValidationSetting.IsRequired (a plain bool) —
// there are no fielded min/max/length/pattern properties anywhere in this schema. Anything
// beyond required is expressed as the ValidWhen conditional formula, captured but unevaluated.
function readValidation(el: Record<string, unknown>): FieldValidation | undefined {
  const validationSetting = gf(el, "validationSetting") as Record<string, unknown> | undefined;
  if (!validationSetting) return undefined;
  const required = !!gf(validationSetting, "isRequired");
  const validWhen = gf(validationSetting, "validWhen");
  const v: FieldValidation = {};
  if (required) v.required = true;
  if (validWhen !== null && validWhen !== undefined) v.formula = validWhen;
  return Object.keys(v).length > 0 ? v : undefined;
}

// BasicSetting.{Label,HelpText,Tooltip} — the template author's own copy for this field.
function readLabelInfo(el: Record<string, unknown>): Pick<FieldMeta, "label" | "helpText" | "tooltip"> | undefined {
  const basic = gf(el, "basicSetting") as Record<string, unknown> | undefined;
  if (!basic) return undefined;
  const label = gf(basic, "label");
  const helpText = gf(basic, "helpText");
  const tooltip = gf(basic, "tooltip");
  const out: Pick<FieldMeta, "label" | "helpText" | "tooltip"> = {};
  if (typeof label === "string" && label) out.label = label;
  if (typeof helpText === "string" && helpText) out.helpText = helpText;
  if (typeof tooltip === "string" && tooltip) out.tooltip = tooltip;
  return Object.keys(out).length > 0 ? out : undefined;
}

// DataSetting.{DefaultValue,DefaultValueType} — only trust it as a prefillable literal when
// DefaultValueType is "Custom"; other types (formula/variable-driven) aren't a value we can
// just drop into the input, so they're skipped rather than guessed at.
function readDefaultValue(el: Record<string, unknown>): unknown {
  const dataSetting = gf(el, "dataSetting") as Record<string, unknown> | undefined;
  if (!dataSetting) return undefined;
  const value = gf(dataSetting, "defaultValue");
  if (value === null || value === undefined || value === "") return undefined;
  const type = gf(dataSetting, "defaultValueType");
  if (typeof type === "string" && type.toLowerCase() !== "custom") return undefined;
  return value;
}

/**
 * Recursively walks ServiceResult.FormElements (nested via ChildElements — sections/pages/
 * containers/fields all share this tree shape) collecting per-field validation + static DOV
 * metadata, keyed by the field's bound variable name (BindVariableFullName; falls back to
 * Name for elements without a variable binding, e.g. table columns).
 */
export function extractFieldMeta(rawEnvelope: Record<string, unknown>): Record<string, FieldMeta> {
  const out: Record<string, FieldMeta> = {};
  // GET /v3/forms/{formId} wraps the actual definition in a ServiceResult envelope
  // alongside Timestamp/RequestId/Error.
  const root = (gf(rawEnvelope, "serviceResult") as Record<string, unknown> | undefined) ?? rawEnvelope;
  const elements = (gf(root, "formElements") as Array<Record<string, unknown>> | undefined) ?? [];

  let visited = 0;
  const visit = (el: Record<string, unknown>) => {
    visited++;
    const name = String(gf(el, "bindVariableFullName") ?? gf(el, "name") ?? "");
    if (name) {
      const validation = readValidation(el);
      const options = readStaticOptions(el);
      const labelInfo = readLabelInfo(el);
      const defaultValue = readDefaultValue(el);
      if (validation || options || labelInfo || defaultValue !== undefined) {
        out[name] = {
          ...(validation ? { validation } : {}),
          ...(options ? { options } : {}),
          ...labelInfo,
          ...(defaultValue !== undefined ? { defaultValue } : {}),
        };
      }
    }
    const children = gf(el, "childElements") as Array<Record<string, unknown>> | undefined;
    (children ?? []).forEach(visit);
  };
  elements.forEach(visit);

  // Table columns bind as "TableName.ColumnName" (confirmed live: a "Choice_list" table's
  // "s" column binds as "Choice_list.s"), but buildFormSchema looks columns up by their bare
  // name. Add an unqualified alias when the suffix is unambiguous across the whole form.
  const suffixCounts = new Map<string, number>();
  for (const key of Object.keys(out)) {
    const dot = key.lastIndexOf(".");
    if (dot === -1) continue;
    const suffix = key.slice(dot + 1);
    suffixCounts.set(suffix, (suffixCounts.get(suffix) ?? 0) + 1);
  }
  for (const key of Object.keys(out)) {
    const dot = key.lastIndexOf(".");
    if (dot === -1) continue;
    const suffix = key.slice(dot + 1);
    if (suffixCounts.get(suffix) === 1 && !(suffix in out)) out[suffix] = out[key];
  }

  log(`extracted meta for ${Object.keys(out).length} field(s) out of ${visited} elements visited (${elements.length} top-level): ${Object.keys(out).join(", ") || "(none)"}`);
  return out;
}

// Inferred from the one live sample seen so far (2026-08-22): Type 2 elements are named
// "Page" and sit directly under a Section, each wrapping one screenful of fields. This is a
// numeric-enum guess, not a documented contract — if a multi-page template ever produces zero
// or a suspicious page count, re-check mcp-livedoc-debug-form-definition.json's Type values.
const PAGE_ELEMENT_TYPE = 2;

export interface PageGroup {
  id: string;
  label: string;
  // Bound variable names found in this page's subtree — includes both the qualified
  // ("Table.Column") and, where unambiguous, bare suffix form, so the frontend can match
  // against either adhocScalars/adhocTables names or raw bind names. Only fields bound to
  // AD_HOC (i.e. adhoc scalars/tables) land here.
  fieldNames: string[];
  // Names of real (non-AD_HOC) variable lists bound anywhere in this page's subtree — a page
  // can be dedicated entirely to one variable list's fields, in which case fieldNames is empty
  // and the frontend must render the matching variableLists[] entry instead.
  variableListNames: string[];
}

/**
 * Groups bound fields by their containing "Page" element, so the App panel can paginate a
 * wizard along the same boundaries the template author actually drew, instead of a generic
 * field-count heuristic. Returns [] when the form has 0 or 1 pages (nothing to page against).
 */
export function extractPageGroups(rawEnvelope: Record<string, unknown>): PageGroup[] {
  const root = (gf(rawEnvelope, "serviceResult") as Record<string, unknown> | undefined) ?? rawEnvelope;
  const elements = (gf(root, "formElements") as Array<Record<string, unknown>> | undefined) ?? [];

  // A field's BindVariableListName is "AD_HOC" for adhoc scalars/tables, or the real variable
  // list's name (e.g. "SQL0526_1") when bound to one. Fields bound to a real list never appear
  // in adhocScalars/adhocTables, so their names must NOT be added to fieldNames — the frontend
  // renders those via variableListNames instead (see FormApp.jsx pageVlIndices).
  const collectBoundNames = (el: Record<string, unknown>, names: Set<string>, listNames: Set<string>) => {
    const bind = gf(el, "bindVariableFullName");
    const listName = gf(el, "bindVariableListName");
    const isRealList = typeof listName === "string" && listName && listName !== "AD_HOC";
    if (isRealList) {
      listNames.add(listName);
    } else if (typeof bind === "string" && bind) {
      names.add(bind);
      const dot = bind.lastIndexOf(".");
      if (dot !== -1) names.add(bind.slice(dot + 1));
    }
    const children = gf(el, "childElements") as Array<Record<string, unknown>> | undefined;
    (children ?? []).forEach(c => collectBoundNames(c, names, listNames));
  };

  const groups: PageGroup[] = [];
  const visit = (el: Record<string, unknown>) => {
    if (Number(gf(el, "type")) === PAGE_ELEMENT_TYPE) {
      const names = new Set<string>();
      const listNames = new Set<string>();
      collectBoundNames(el, names, listNames);
      if (names.size > 0 || listNames.size > 0) {
        const basic = gf(el, "basicSetting") as Record<string, unknown> | undefined;
        const label = String((basic && gf(basic, "label")) ?? gf(el, "name") ?? `Page ${groups.length + 1}`);
        groups.push({ id: String(gf(el, "id") ?? `page-${groups.length}`), label, fieldNames: [...names], variableListNames: [...listNames] });
      }
      return; // don't recurse into a page's own children looking for nested pages
    }
    const children = gf(el, "childElements") as Array<Record<string, unknown>> | undefined;
    (children ?? []).forEach(visit);
  };
  elements.forEach(visit);

  log(`found ${groups.length} page group(s): ${groups.map(g => `${g.label}(${g.fieldNames.length} field(s), lists=[${g.variableListNames.join(",")}])`).join(", ") || "(none)"}`);
  return groups.length > 1 ? groups : [];
}

/** Dumps the raw form-definition response to a temp file for shape inspection (best-effort). */
export function dumpFormDefinitionDebug(raw: Record<string, unknown>): void {
  try {
    fs.writeFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug-form-definition.json"), JSON.stringify(raw, null, 2), "utf-8");
  } catch { /* best-effort debug dump */ }
}

export async function handleGetFormDefinition(args: { formId: string; teamSiteId: string }): Promise<
  { formId: string; teamSiteId: string; fields: Array<{ name: string } & FieldMeta>; pageGroups: PageGroup[] }
  | { error: string; detail: unknown }
> {
  const result = await fetchFormDefinition(args.formId, args.teamSiteId);
  if (result.status !== 200) {
    return { error: `Failed to get form definition (HTTP ${result.status})`, detail: result.body };
  }
  const raw = result.body as Record<string, unknown>;
  // Dump raw response so the actual FormElements/validation/DOV shape can be inspected —
  // same debug convention as mcp-livedoc-debug-forms.json in inputs.ts.
  dumpFormDefinitionDebug(raw);

  const meta = extractFieldMeta(raw);
  return {
    formId: args.formId,
    teamSiteId: args.teamSiteId,
    fields: Object.entries(meta).map(([name, m]) => ({ name, ...m })),
    pageGroups: extractPageGroups(raw),
  };
}
