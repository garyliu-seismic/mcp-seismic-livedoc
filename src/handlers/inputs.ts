import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { apiFetch, gf, isComplex } from "../api/client.js";
import { needsContentResolution, resolveManualSelectCandidates } from "./content.js";
import { fetchFormDefinition, extractFieldMeta, extractPageGroups, dumpFormDefinitionDebug, FieldMeta, PageGroup } from "./formDefinition.js";
import { DEBUG_LOG } from "../utils/debug.js";

function logFormDef(msg: string): void {
  try { fs.appendFileSync(DEBUG_LOG, `[${new Date().toISOString()}] FORMDEF: ${msg}\n`); } catch { /* best-effort */ }
}

export async function handleGetInputs(args: {
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
  fieldMeta: Record<string, FieldMeta>;
  pageGroups: PageGroup[];
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
  const msItemsPromise = Promise.all(
    msItems.map(async (item) => {
      const contentType = String(gf(item, "contentType") ?? "");
      if (!needsContentResolution(contentType)) return;
      const resolved = await resolveManualSelectCandidates(item);
      item.candidates = resolved.candidates;
      item.candidatesTotalCount = resolved.totalCount;
    })
  );

  // Best-effort: pull validation rules + static domain-of-value lists from the form
  // definition (GET /v3/forms/{formId}) for the default form, and merge them into the
  // per-field schema below. Never fatal — the higher-level input schema above is still
  // usable without it. Kicked off in parallel with msItemsPromise above (rather than after
  // it) since this call has been observed to take the full 30s timeout on a slow QA
  // environment — running it sequentially would add that on top of everything else.
  let fieldMeta: Record<string, FieldMeta> = {};
  let pageGroups: PageGroup[] = [];
  const defaultForm = forms.find(f => gf(f, "isDefault")) ?? forms[0];
  const formId = defaultForm ? String(gf(defaultForm, "id") ?? "") : "";
  const formDefPromise = !formId
    ? Promise.resolve(logFormDef(`handleGetInputs: no formId available (forms.length=${forms.length}) — skipping validation/DOV lookup`))
    : fetchFormDefinition(formId, args.teamSiteId).then(defRes => {
        if (defRes.status === 200) {
          const rawDef = defRes.body as Record<string, unknown>;
          dumpFormDefinitionDebug(rawDef);
          fieldMeta = extractFieldMeta(rawDef);
          pageGroups = extractPageGroups(rawDef);
          logFormDef(`handleGetInputs: merged ${Object.keys(fieldMeta).length} field(s) of validation/DOV metadata, ${pageGroups.length} page group(s) for formId=${formId}`);
        } else {
          logFormDef(`handleGetInputs: form definition lookup FAILED for formId=${formId} — proceeding without validation/DOV metadata`);
        }
      }).catch(e => {
        logFormDef(`handleGetInputs: form definition lookup THREW for formId=${formId}: ${e}`);
      });

  await Promise.all([msItemsPromise, formDefPromise]);

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
    fieldMeta,
    pageGroups,
  };
}

export function buildFormSchema(
  ir: {
    templateName: string;
    teamSiteId: string;
    libraryContentVersionId: string;
    adhocInputs: Array<Record<string, unknown>>;
    variableListData: Array<Record<string, unknown>>;
    manualSelectContentInput: Record<string, unknown> | undefined;
    forms: Array<Record<string, unknown>>;
    fieldMeta?: Record<string, FieldMeta>;
    pageGroups?: PageGroup[];
  },
  token: string
): unknown {
  const isTable = (i: Record<string, unknown>) => ((gf(i, "columns") as unknown[] | undefined)?.length ?? 0) > 0;
  const fieldMeta = ir.fieldMeta ?? {};
  // Field names in adhocInputs/variableListData don't reliably match the case used in the
  // form definition, so look up case-insensitively.
  const metaByLowerName = new Map(Object.entries(fieldMeta).map(([k, v]) => [k.toLowerCase(), v]));
  const metaFor = (name: string): FieldMeta => metaByLowerName.get(name.toLowerCase()) ?? {};

  const adhocScalars = ir.adhocInputs.filter(i => !isTable(i)).map(i => {
    const name = String(gf(i, "name") ?? "");
    return { name, type: String(gf(i, "type") ?? "STRING"), ...metaFor(name) };
  });

  const adhocTables = ir.adhocInputs.filter(i => isTable(i)).map(i => ({
    name: String(gf(i, "name") ?? ""),
    columns: ((gf(i, "columns") as Array<Record<string, unknown>>) ?? []).map(c => {
      const name = String(gf(c, "name") ?? "");
      return { name, colType: String(gf(c, "colType") ?? gf(c, "type") ?? "TEXT"), ...metaFor(name) };
    }),
  }));

  const variableLists = ir.variableListData.map(vl => {
    const inputs = (gf(vl, "variableInputs") as Array<Record<string, unknown>>) ?? [];
    return {
      name: String(gf(vl, "variableListName") ?? ""),
      dataSourceName: String(gf(vl, "dataSourceName") ?? gf(vl, "dataSourceId") ?? ""),
      scalars: inputs.filter(i => !isTable(i)).map(i => {
        const name = String(gf(i, "name") ?? "");
        return { name, type: String(gf(i, "type") ?? "STRING"), ...metaFor(name) };
      }),
      tables: inputs.filter(i => isTable(i)).map(i => ({
        name: String(gf(i, "name") ?? ""),
        columns: ((gf(i, "columns") as Array<Record<string, unknown>>) ?? []).map(c => {
          const name = String(gf(c, "name") ?? "");
          return { name, colType: String(gf(c, "colType") ?? gf(c, "type") ?? "TEXT"), ...metaFor(name) };
        }),
      })),
    };
  });

  const msItems = (ir.manualSelectContentInput
    ? gf(ir.manualSelectContentInput, "manualSelectContentItems")
    : undefined) as Array<Record<string, unknown>> | undefined ?? [];

  const slideGroups = msItems
    .filter(i => ["Group", "Section"].includes(String(gf(i, "contentType") ?? "")))
    .map(i => ({
      id: String(gf(i, "id") ?? ""),
      name: String(gf(i, "name") ?? ""),
      contentType: String(gf(i, "contentType") ?? "Group"),
      orderIndex: Number(gf(i, "orderIndex") ?? 0),
      defaultInclude: gf(i, "isInclude") !== false,
      thumbnailUrl: String(
        gf(i, "thumbnailUrl") ?? gf(i, "previewImageUrl") ?? gf(i, "imageUrl") ??
        ((gf(i, "pages") ?? gf(i, "Pages")) as Array<Record<string, unknown>> | undefined)?.[0]?.ImageUrl ??
        ((gf(i, "pages") ?? gf(i, "Pages")) as Array<Record<string, unknown>> | undefined)?.[0]?.imageUrl ??
        ""
      ),
    }));

  const externalContent = msItems
    .filter(i => !["Group", "Section"].includes(String(gf(i, "contentType") ?? "")))
    .map(i => ({
      id: String(gf(i, "id") ?? ""),
      name: String(gf(i, "name") ?? ""),
      contentType: String(gf(i, "contentType") ?? ""),
      orderIndex: Number(gf(i, "orderIndex") ?? 0),
      candidates: ((i.candidates as Array<Record<string, unknown>>) ?? []).map(c => ({
        versionId: String(c.versionId ?? ""),
        contentId: String(c.contentId ?? ""),
        ...(c.sourceBlobId ? { sourceBlobId: String(c.sourceBlobId) } : {}),
        title: String(c.title ?? ""),
        format: String(c.format ?? "PPTX"),
        thumbnailUrl: String(c.thumbnailUrl ?? ""),
      })),
    }));

  // Dump raw forms to a temp debug file so we can inspect the actual API shape.
  fs.writeFileSync(path.join(os.tmpdir(), `mcp-livedoc-debug-forms.json`), JSON.stringify(ir.forms, null, 2), "utf-8");

  // Group forms by name; collect distinct output combinations per group.
  const formsByName = new Map<string, { id: string; combos: Array<Array<{ format: string; name?: string }>> }>();
  for (const f of ir.forms) {
    const name = String(gf(f, "name") ?? "");
    const id = String(gf(f, "id") ?? "");
    const rawOutputs = (gf(f, "outputs") ?? gf(f, "outputFormats") ?? gf(f, "outputDefinitions")) as Array<Record<string, unknown>> | undefined;
    const outputs = (rawOutputs ?? []).map(o => ({
      format: String(gf(o, "format") ?? gf(o, "outputFormat") ?? "").toUpperCase(),
      name: String(gf(o, "name") ?? gf(o, "displayName") ?? ""),
    })).filter(o => o.format);
    if (!formsByName.has(name)) formsByName.set(name, { id, combos: [] });
    formsByName.get(name)!.combos.push(outputs);
  }
  const formOptions = Array.from(formsByName.entries()).map(([name, { id, combos }]) => ({ name, id, outputCombos: combos }));

  return {
    token, templateName: ir.templateName, teamSiteId: ir.teamSiteId, libraryContentVersionId: ir.libraryContentVersionId,
    adhocScalars, adhocTables, variableLists, slideGroups, externalContent, formOptions,
    pageGroups: ir.pageGroups ?? [],
  };
}
