import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { apiFetch, gf, isComplex } from "../api/client.js";
import { needsContentResolution, resolveManualSelectCandidates } from "./content.js";

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

export function buildFormSchema(
  ir: {
    templateName: string;
    teamSiteId: string;
    libraryContentVersionId: string;
    adhocInputs: Array<Record<string, unknown>>;
    variableListData: Array<Record<string, unknown>>;
    manualSelectContentInput: Record<string, unknown> | undefined;
    forms: Array<Record<string, unknown>>;
  },
  token: string
): unknown {
  const isTable = (i: Record<string, unknown>) => ((gf(i, "columns") as unknown[] | undefined)?.length ?? 0) > 0;

  const adhocScalars = ir.adhocInputs.filter(i => !isTable(i)).map(i => ({
    name: String(gf(i, "name") ?? ""),
    type: String(gf(i, "type") ?? "STRING"),
  }));

  const adhocTables = ir.adhocInputs.filter(i => isTable(i)).map(i => ({
    name: String(gf(i, "name") ?? ""),
    columns: ((gf(i, "columns") as Array<Record<string, unknown>>) ?? []).map(c => ({
      name: String(gf(c, "name") ?? ""),
      colType: String(gf(c, "colType") ?? gf(c, "type") ?? "TEXT"),
    })),
  }));

  const variableLists = ir.variableListData.map(vl => {
    const inputs = (gf(vl, "variableInputs") as Array<Record<string, unknown>>) ?? [];
    return {
      name: String(gf(vl, "variableListName") ?? ""),
      dataSourceName: String(gf(vl, "dataSourceName") ?? gf(vl, "dataSourceId") ?? ""),
      scalars: inputs.filter(i => !isTable(i)).map(i => ({
        name: String(gf(i, "name") ?? ""),
        type: String(gf(i, "type") ?? "STRING"),
      })),
      tables: inputs.filter(i => isTable(i)).map(i => ({
        name: String(gf(i, "name") ?? ""),
        columns: ((gf(i, "columns") as Array<Record<string, unknown>>) ?? []).map(c => ({
          name: String(gf(c, "name") ?? ""),
          colType: String(gf(c, "colType") ?? gf(c, "type") ?? "TEXT"),
        })),
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
  const formsByName = new Map<string, Array<Array<{ format: string; name?: string }>>>();
  for (const f of ir.forms) {
    const name = String(gf(f, "name") ?? "");
    const rawOutputs = (gf(f, "outputs") ?? gf(f, "outputFormats") ?? gf(f, "outputDefinitions")) as Array<Record<string, unknown>> | undefined;
    const outputs = (rawOutputs ?? []).map(o => ({
      format: String(gf(o, "format") ?? gf(o, "outputFormat") ?? "").toUpperCase(),
      name: String(gf(o, "name") ?? gf(o, "displayName") ?? ""),
    })).filter(o => o.format);
    if (!formsByName.has(name)) formsByName.set(name, []);
    formsByName.get(name)!.push(outputs);
  }
  const formOptions = Array.from(formsByName.entries()).map(([name, combos]) => ({ name, outputCombos: combos }));

  return { token, templateName: ir.templateName, teamSiteId: ir.teamSiteId, libraryContentVersionId: ir.libraryContentVersionId, adhocScalars, adhocTables, variableLists, slideGroups, externalContent, formOptions };
}
