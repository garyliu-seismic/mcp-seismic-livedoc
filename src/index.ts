#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerAppTool, registerAppResource, RESOURCE_MIME_TYPE, getUiCapability } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import { randomUUID } from "crypto";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { exec } from "child_process";
import { fileURLToPath } from "url";


// Reads coworkUserFilesPath from the Claude Desktop config JSON.
// Tries the standard %APPDATA%\Claude path first, then the Microsoft Store
// package path (%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude).
function findCoworkPath(): string {
  const candidates: string[] = [];
  const appdata = process.env.APPDATA;
  const localAppdata = process.env.LOCALAPPDATA;
  if (appdata) candidates.push(path.join(appdata, "Claude", "claude_desktop_config.json"));
  if (localAppdata) {
    const pkgsDir = path.join(localAppdata, "Packages");
    try {
      for (const entry of fs.readdirSync(pkgsDir)) {
        if (entry.startsWith("Claude_")) {
          candidates.push(path.join(pkgsDir, entry, "LocalCache", "Roaming", "Claude", "claude_desktop_config.json"));
        }
      }
    } catch { /* Packages dir unreadable */ }
  }
  for (const cfgPath of candidates) {
    try {
      const cfg = JSON.parse(fs.readFileSync(cfgPath, "utf8")) as Record<string, unknown>;
      const p = cfg.coworkUserFilesPath;
      if (typeof p === "string" && p) return p;
    } catch { /* not found or not parseable */ }
  }
  return "";
}

const COWORK_PATH = process.env.CLAUDE_COWORK_PATH || findCoworkPath();

const BASE_URL = process.env.SEISMIC_BASE_URL ?? "https://api.seismic.com/livedoc";

const DEFAULT_AUTH_URI      = process.env.AUTH_SERVICE_URI    ?? "";
const DEFAULT_AUTH_TENANT   = process.env.AUTH_TENANT         ?? "";
const DEFAULT_CLIENT_ID     = process.env.AUTH_CLIENT_ID      ?? "";
const DEFAULT_CLIENT_SECRET = process.env.AUTH_CLIENT_SECRET  ?? "";
const DEFAULT_USERNAME      = process.env.AUTH_USERNAME        ?? "";
const DEFAULT_PASSWORD      = process.env.AUTH_PASSWORD        ?? "";

const pendingFormSchemas  = new Map<string, unknown>(); // token → normalised form schema for get_form_schema
const pendingGenerations  = new Map<string, string>();  // generatedLivedocId → formToken (Process 2 only)

let currentToken = process.env.SEISMIC_API_TOKEN ?? "";
// True once a token was explicitly provided via set_token — disables the silent
// 401-triggered autoLogin() so it can never clobber a hand-picked token with a
// narrower-scoped one obtained from the default credential-flow login.
let tokenIsManual = false;

// Scopes requested by the credential-flow login (login tool + autoLogin refresh).
// Must include seismic.library.view/manage — search endpoints reject tokens
// without them ("Invalid or missing user claim"), even though generation
// endpoints are fine with just livedoc/library.
const LOGIN_SCOPE = "library livedoc seismic.library.view seismic.library.manage";

function authHeaders(): Record<string, string> {
  return {
    Authorization: `Bearer ${currentToken}`,
    "Content-Type": "application/json",
  };
}

async function autoLogin(): Promise<boolean> {
  if (!DEFAULT_AUTH_URI || !DEFAULT_AUTH_TENANT || !DEFAULT_CLIENT_ID || !DEFAULT_USERNAME || !DEFAULT_PASSWORD) return false;
  try {
    const body = new URLSearchParams({
      grant_type:    "client_credentials",
      client_id:     DEFAULT_CLIENT_ID,
      client_secret: DEFAULT_CLIENT_SECRET,
      username:      DEFAULT_USERNAME,
      password:      DEFAULT_PASSWORD,
      scope:         LOGIN_SCOPE,
    });
    const res = await fetch(`${DEFAULT_AUTH_URI}/tenants/${encodeURIComponent(DEFAULT_AUTH_TENANT)}/connect/token`, {
      method:  "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body:    body.toString(),
    });
    if (!res.ok) return false;
    const data = await res.json() as Record<string, unknown>;
    if (!data.access_token) return false;
    currentToken = data.access_token as string;
    tokenIsManual = false;
    return true;
  } catch {
    return false;
  }
}

async function apiFetch(
  path: string,
  options: RequestInit = {},
  _retry = true
): Promise<{ status: number; body: unknown }> {
  const url = `${BASE_URL}${path}`;
  const res = await fetch(url, {
    ...options,
    headers: { ...authHeaders(), ...(options.headers as Record<string, string> ?? {}) },
  });
  // Auto-refresh token on 401 when env credentials are available — but never when the
  // current token was explicitly set via set_token, so we don't silently replace a
  // hand-picked (possibly broader-scoped) token with a narrower credential-flow one.
  if (res.status === 401 && _retry && !tokenIsManual && DEFAULT_USERNAME && DEFAULT_PASSWORD) {
    const refreshed = await autoLogin();
    if (refreshed) return apiFetch(path, options, false);
  }
  const text = await res.text();
  let body: unknown;
  try {
    body = JSON.parse(text);
  } catch {
    body = text;
  }
  return { status: res.status, body };
}

const FORM_APP_BASE = process.env.FORM_APP_URL ?? "http://localhost:5173";
const FORM_API_BASE = process.env.FORM_API_URL ?? "http://localhost:3001";

function isComplex(resp: Record<string, unknown>): boolean {
  const vlData = (resp.variableListData ?? resp.VariableListData) as Array<Record<string, unknown>> | undefined;
  const imgInput = (resp.imageUploadContentInput ?? resp.ImageUploadContentInput) as Record<string, unknown> | undefined;
  const manualSelect = (resp.manualSelectContentInput ?? resp.ManualSelectContentInput) as Record<string, unknown> | undefined;
  const adhoc = (resp.adhocInputs ?? resp.AdhocInputs) as Array<unknown> | undefined;
  return !!(
    vlData?.some(v => v.dataSourceId ?? v.DataSourceId) ||
    (imgInput?.imageUploadContentItems as Array<unknown> | undefined)?.length ||
    (manualSelect?.manualSelectContentItems as Array<unknown> | undefined)?.length ||
    (adhoc?.length ?? 0) > 8
  );
}

// ── Tool definitions ────────────────────────────────────────────────────────

const FORM_RESOURCE_URI = "ui://livedoc/form";

// ── Tool handlers ───────────────────────────────────────────────────────────

async function handleSearchTemplates(args: {
  searchText?: string;
  page_size?: number;
}) {
  const size = Math.min(args.page_size ?? 10, 50);
  const body = {
    searchText: args.searchText ?? "",
    allowPptx: true,
    includeLiveDoc: true,
    allowPdf: false,
    page: { size, from: 0 },
    orderBy: [{ attr: "modifiedDate", direction: "DESC" }],
  };
  const result = await apiFetch("/v3/contents", {
    method: "POST",
    body: JSON.stringify(body),
  });
  if (result.status !== 200) {
    const detail = result.body as Record<string, unknown> | undefined;
    const isUserClaimError = typeof detail === "object" && String(detail?.Message ?? "").includes("user claim");
    if (isUserClaimError) {
      return {
        error: "search_requires_user_token",
        message: "Template search requires a user-context token. The current token is a service account token without user identity claims. To enable search: set SEISMIC_API_TOKEN in claude_desktop_config.json to a user token (obtain one from the Seismic dev portal or browser DevTools). If you already know your template's teamSiteId and libraryContentVersionId, call get_livedoc_inputs directly — authentication for generation is not affected.",
      };
    }
    return { error: `Search failed (HTTP ${result.status})`, detail: result.body };
  }
  const data = result.body as {
    totalCount: number;
    documents: Array<{
      contentId: string;
      contentVersionId: string;
      title: string;
      description: string | null;
      format: string;
      teamsite: string;
      modifiedDate: string;
    }>;
  };
  return {
    totalCount: data.totalCount,
    results: data.documents.map((d) => ({
      title: d.title,
      format: d.format,
      contentVersionId: d.contentVersionId,
      teamSiteId: d.teamsite,
      modifiedDate: d.modifiedDate,
      description: d.description,
    })),
  };
}

async function handleSearchContent(args: {
  query: string;
  contentType?: string;
  page_size?: number;
}) {
  const size = Math.min(args.page_size ?? 10, 50);
  // Map contentType to /v3/contents flags. allowPptx requires at least one of
  // includeStandardPptx/includeLiveDoc also true, or the API rejects the request.
  const ct = args.contentType ?? "";
  const isSlideType = !ct || ["ExternalSlides", "LiveSlide", "ExternalStaticSlides"].includes(ct);
  const isLiveDocType = !ct || ct === "LiveDoc";
  const allowPptx = isSlideType || isLiveDocType;
  const includeStandardPptx = isSlideType;
  const includeLiveDoc = isSlideType || isLiveDocType;
  const allowPdf = !ct || ct === "PDF";
  const body = {
    searchText: args.query,
    allowPptx,
    includeStandardPptx,
    includeLiveDoc,
    allowPdf,
    page: { size, from: 0 },
    orderBy: [{ attr: "modifiedDate", direction: "DESC" }],
  };
  const result = await apiFetch("/v3/contents", {
    method: "POST",
    body: JSON.stringify(body),
  });
  if (result.status !== 200) {
    return { error: `Search failed (HTTP ${result.status})`, detail: result.body };
  }
  const data = result.body as {
    totalCount: number;
    documents: Array<{
      contentId: string;
      contentVersionId: string;
      title: string;
      format: string;
      teamsite: string;
      modifiedDate: string;
      sourceBlobId?: string;
    }>;
  };
  return {
    totalCount: data.totalCount,
    results: data.documents.map((d) => ({
      id: d.contentVersionId,
      name: d.title,
      format: d.format,
      contentVersionId: d.contentVersionId,
      sourceBlobId: d.sourceBlobId,
      modifiedDate: d.modifiedDate,
    })),
  };
}

// "Group"/"Section" are the only manualSelectContentItem types that are already fully valid
// as returned — every other contentType needs real content resolved via search before submission.
// This is a denylist rather than an allowlist because the GET side's vocabulary doesn't match
// the submission-side ManualSelectContentType enum 1:1 (e.g. GET can return "ExternalSlides",
// which isn't even a valid value to submit — it must be resolved then re-mapped to "LiveSlide"
// or "ResourcePDF" depending on the chosen candidate's format).
function needsContentResolution(contentType: string): boolean {
  return contentType !== "" && contentType !== "Group" && contentType !== "Section";
}

// C# bool property names here don't follow simple camelCase (AllowPDF, IncludeStandardPPTX),
// so check several literal casings rather than relying on gf()'s single-fallback capitalization.
function boolField(item: Record<string, unknown>, ...keys: string[]): boolean {
  for (const k of keys) {
    if (typeof item[k] === "boolean") return item[k] as boolean;
  }
  return false;
}

// Resolves real content candidates for one manualSelectContentItem. Prefers the item's own
// filter/format flags (the template author's actual search criteria, e.g. Filter: [{propertyName:
// "ContentName", operator: "CT", value: "sp3"}]) over a generic name-based guess — those flags
// are what get_livedoc_inputs actually returns on ExternalSlideContent items.
const CANDIDATE_PAGE_SIZE = 10;

async function resolveManualSelectCandidates(item: Record<string, unknown>): Promise<{ candidates: Array<Record<string, unknown>>; totalCount: number }> {
  const name = String(gf(item, "name") ?? "");
  const contentType = String(gf(item, "contentType") ?? "");
  const filter = (gf(item, "filter") as unknown[] | undefined) ?? [];
  const rawIsApplyAllFilter = gf(item, "isApplyAllFilter");
  const isApplyAllFilter = typeof rawIsApplyAllFilter === "boolean" ? rawIsApplyAllFilter : true;

  let allowPptx = boolField(item, "allowPptx", "AllowPptx");
  let includeStandardPptx = boolField(item, "includeStandardPptx", "IncludeStandardPPTX", "IncludeStandardPptx");
  let includeLiveDoc = boolField(item, "includeLiveDoc", "IncludeLiveDoc");
  let allowPdf = boolField(item, "allowPdf", "AllowPDF", "AllowPdf");

  // Fall back to a contentType-based guess only if the item carried no usable format flags at all.
  if (!allowPptx && !allowPdf) {
    const isSlideType = ["ExternalSlides", "LiveSlide", "ExternalStaticSlides"].includes(contentType);
    allowPptx = isSlideType;
    includeStandardPptx = isSlideType;
    includeLiveDoc = isSlideType;
    allowPdf = !isSlideType;
  }

  const body: Record<string, unknown> = {
    allowPptx,
    includeStandardPptx,
    includeLiveDoc,
    allowPdf,
    page: { size: CANDIDATE_PAGE_SIZE, from: 0 },
    orderBy: [{ attr: "modifiedDate", direction: "DESC" }],
  };
  if (filter.length > 0) {
    // The item's own filter is the template author's actual search criteria — combining it
    // with a searchText:name guess (name is just a display label, e.g. "sp3") over-constrains
    // the query and silently returns zero results, so filter and searchText are mutually exclusive here.
    body.filter = filter;
    body.isApplyAllFilter = isApplyAllFilter;
  } else {
    body.searchText = name;
  }

  const result = await apiFetch("/v3/contents", { method: "POST", body: JSON.stringify(body) });
  if (result.status !== 200) return { candidates: [], totalCount: 0 };
  const data = result.body as { documents?: Array<Record<string, unknown>>; totalCount?: number };
  const candidates = (data.documents ?? []).slice(0, CANDIDATE_PAGE_SIZE).map((d) => ({
    versionId: gf(d, "contentVersionId"),
    contentId: gf(d, "contentId"),
    sourceBlobId: gf(d, "sourceBlobId"),
    title: gf(d, "title"),
    format: gf(d, "format"),
    thumbnailUrl: String(gf(d, "thumbnailUrl") ?? ""),
  }));
  return { candidates, totalCount: data.totalCount ?? candidates.length };
}

async function handleGetInputs(args: {
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

function gf(o: Record<string, unknown>, key: string): unknown {
  return o[key] ?? o[key[0].toUpperCase() + key.slice(1)];
}

function buildFormSchema(
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

async function handleSubmitGeneration(args: {
  teamSiteId: string;
  libraryContentVersionId: string;
  adHocInputs: Array<{ name: string; value: unknown }>;
  outputs: Array<{ format: string; name?: string; fileName?: string }>;
  variableListData?: Array<{
    variableListName: string;
    variableInputs: Array<{ name: string; value: unknown }>;
  }>;
  liveFormSellerTemplateId?: string;
  regionalFormat?: string;
  manualSelectContentInput?: {
    manualSelectContentItems: Array<{
      id: string;
      name?: string;
      contentType: string;
      versionId?: string;
      sourceBlobId?: string;
      pageNumber?: number;
      isInclude: boolean;
      orderIndex?: number;
    }>;
  };
}) {
  if (args.manualSelectContentInput) {
    const unresolved = args.manualSelectContentInput.manualSelectContentItems.filter(
      (item) =>
        item.isInclude !== false &&
        needsContentResolution(item.contentType) &&
        !item.versionId
    );
    if (unresolved.length > 0) {
      return {
        error: "WRONG TOOL — do not call submit_livedoc_generation directly when manualSelectContentInput has unresolved items.",
        detail: `Item(s) [${unresolved.map((i) => `"${i.name ?? i.id}"`).join(", ")}] are missing versionId. You must NOT resolve versionId yourself via search_livedoc_content or any other tool. The correct flow is: (1) get_livedoc_inputs opens the form in the App panel, (2) the USER fills it out and clicks Submit — the payload is copied to their clipboard, (3) the user pastes the payload into the chat, (4) THEN call this tool with that exact pasted JSON.`,
      };
    }
  }

  const reqBody: Record<string, unknown> = {
    adHocInputs: args.adHocInputs,
    outputs: args.outputs,
  };
  if (args.variableListData) reqBody.variableListData = args.variableListData;
  if (args.regionalFormat) reqBody.regionalFormat = args.regionalFormat;
  if (args.manualSelectContentInput) {
    reqBody.manualSelectContentInput = {
      manualSelectContentItems: args.manualSelectContentInput.manualSelectContentItems.map((item) => ({
        id: item.id,
        name: item.name,
        contentType: item.contentType,
        versionId: item.versionId,
        sourceBlobId: item.sourceBlobId,
        pageNumber: item.pageNumber,
        isInclude: item.isInclude,
        orderIndex: item.orderIndex,
      })),
    };
  }

  const qp = args.liveFormSellerTemplateId
    ? `?liveFormSellerTemplateId=${encodeURIComponent(args.liveFormSellerTemplateId)}`
    : "";

  const result = await apiFetch(
    `/v3/teamsites/${args.teamSiteId}/livedocVersions/${args.libraryContentVersionId}${qp}`,
    { method: "POST", body: JSON.stringify(reqBody) }
  );
  if (result.status !== 201 && result.status !== 200) {
    return { error: `Generation submission failed (HTTP ${result.status})`, detail: result.body };
  }
  const body = result.body as Record<string, unknown>;
  // Walk every key looking for something that looks like a generated-livedoc UUID
  const generatedLivedocId = (
    body.generatedLivedocId ?? body.GeneratedLivedocId ??
    body.id ?? body.Id ??
    body.generatedId ?? body.GeneratedId ??
    body.livedocId ?? body.LivedocId ??
    // last resort: find any string value that looks like a UUID
    Object.values(body).find(v => typeof v === "string" && /^[0-9a-f-]{36}$/i.test(v))
  ) as string | undefined;
  return {
    generatedLivedocId,
    rawBody: body,
    message: "Generation job submitted. Call get_generation_status to poll for completion.",
  };
}

// Matches LiveDocGenStatusResp in app-livedoc-service (PublicAPIV3Controller.ResultStatus.cs):
// Queued=0, Generating=1, Completed=2, Failed=3. The API returns this as a number, not a string,
// so callers must map it before comparing against status names.
const STATUS_NAMES = ["Queued", "Generating", "Completed", "Failed"];

function statusName(raw: unknown): string {
  if (typeof raw === "number" && STATUS_NAMES[raw] !== undefined) {
    return STATUS_NAMES[raw];
  }
  if (typeof raw === "string" && STATUS_NAMES.includes(raw)) {
    return raw;
  }
  return String(raw);
}

async function handleGetStatus(args: { generatedLivedocId: string }) {
  const result = await apiFetch(`/v3/generatedLivedocs/${args.generatedLivedocId}`);
  if (result.status !== 200) {
    return { error: `Status check failed (HTTP ${result.status})`, detail: result.body };
  }
  const raw = result.body as Record<string, unknown>;
  const id = (raw.id ?? raw.Id ?? raw.generatedLivedocId ?? raw.GeneratedLivedocId) as string;
  const rawOutputs = (raw.outputs ?? raw.Outputs ?? []) as Array<Record<string, unknown>>;
  const outputs = rawOutputs.map((o) => ({
    id: (o.id ?? o.Id) as string,
    status: statusName(o.status ?? o.Status),
    format: (o.format ?? o.Format) as string,
    name: (o.name ?? o.Name) as string,
    fileName: (o.fileName ?? o.FileName) as string,
    errorString: (o.errorString ?? o.ErrorString ?? null) as string | null,
  }));
  const allDone = outputs.every((o) => o.status === "Completed" || o.status === "Failed");
  return {
    generatedLivedocId: id,
    allDone,
    outputs,
    hint: allDone
      ? "All outputs done. Call get_generation_download_url with each outputId."
      : "Still generating. Poll again in a few seconds.",
  };
}

async function handleGetDownloadUrl(args: {
  generatedLivedocId: string;
  outputId: string;
}) {
  // Primary: redirect=false returns JSON with downloadUrl
  const result = await apiFetch(
    `/v3/generatedLivedocs/${args.generatedLivedocId}/outputs/${args.outputId}/content?redirect=false`
  );
  if (result.status === 200) {
    return result.body;
  }
  // Fallback: capture the 302 Location header (works when redirect=false returns 403)
  try {
    const res = await fetch(
      `${BASE_URL}/v3/generatedLivedocs/${args.generatedLivedocId}/outputs/${args.outputId}/content`,
      { headers: authHeaders() as Record<string, string>, redirect: "manual" as RequestRedirect }
    );
    const location = res.headers.get("location");
    if (location) return { downloadUrl: location };
  } catch { /* ignore, fall through */ }
  return { error: `Download URL fetch failed (HTTP ${result.status})`, detail: result.body };
}

function getDownloadsDir(): string {
  const dir = path.join(os.homedir(), "Downloads");
  try {
    fs.mkdirSync(dir, { recursive: true });
    return dir;
  } catch {
    return os.tmpdir();
  }
}

// Opens a local file with the OS-registered default application (e.g. double-click behavior).
function openWithDefaultApp(filePath: string) {
  const platform = process.platform;
  const quoted = `"${filePath}"`;
  const cmd =
    platform === "win32" ? `start "" ${quoted}` :
    platform === "darwin" ? `open ${quoted}` :
    `xdg-open ${quoted}`;
  exec(cmd, () => { /* best-effort; failures are non-fatal */ });
}

function uniqueFilePath(dir: string, fileName: string): string {
  const ext = path.extname(fileName);
  const base = path.basename(fileName, ext);
  let candidate = path.join(dir, fileName);
  let n = 1;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${base} (${n})${ext}`);
    n++;
  }
  return candidate;
}

async function handleDownloadGenerationOutput(args: {
  generatedLivedocId: string;
  outputId: string;
  autoOpen?: boolean;
}) {
  const dl = await handleGetDownloadUrl({ generatedLivedocId: args.generatedLivedocId, outputId: args.outputId });
  if (dl && typeof dl === "object" && "error" in (dl as object)) {
    return dl;
  }
  const body = dl as Record<string, unknown>;
  // DownloadLocationResp only ever contains `downloadUrl` — it never carries a fileName.
  // The real fileName (with extension) lives on get_generation_status's outputs[], so look it up there.
  const url = String(body.url ?? body.downloadUrl ?? body.Url ?? body.DownloadUrl ?? "");
  if (!url) {
    return { error: "No download URL returned for this output.", detail: body };
  }

  let fileName = `livedoc-${args.outputId}`;
  const status = await handleGetStatus({ generatedLivedocId: args.generatedLivedocId });
  if (status && typeof status === "object" && "outputs" in (status as object)) {
    const outputs = (status as { outputs: Array<{ id: string; format: string; fileName: string }> }).outputs;
    const match = outputs.find(
      (o) => o.id === args.outputId || (o.format ?? "").toLowerCase() === args.outputId.toLowerCase()
    );
    if (match) {
      const rawName = match.fileName || fileName;
      const ext = path.extname(rawName);
      fileName = ext ? rawName : `${rawName}.${(match.format ?? "pptx").toLowerCase()}`;
    }
  }

  const fileRes = await fetch(url);
  if (!fileRes.ok) {
    return { error: `Failed to download file content (HTTP ${fileRes.status})` };
  }
  const buf = Buffer.from(await fileRes.arrayBuffer());

  const dir = getDownloadsDir();
  const filePath = uniqueFilePath(dir, fileName);
  fs.writeFileSync(filePath, buf);

  const autoOpen = args.autoOpen !== false;
  if (autoOpen) openWithDefaultApp(filePath);

  return {
    filePath,
    fileName,
    sizeBytes: buf.length,
    opened: autoOpen,
    message: autoOpen
      ? `Downloaded "${fileName}" (${buf.length} bytes) to ${filePath} and opened it with the default app.`
      : `Downloaded "${fileName}" (${buf.length} bytes) to ${filePath}.`,
  };
}

function generateToken(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}


async function handleOpenFormUi(args: { teamSiteId: string; libraryContentVersionId: string; context?: string; prefillValues?: unknown }) {
  // Push the current token to the form server so it never uses a stale value.
  try {
    await fetch(`${FORM_API_BASE}/api/set-token`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ token: currentToken }),
    });
  } catch {
    // Non-fatal — form server may not be running yet; it will fall back to its own env var.
  }

  const token = generateToken();
  const params = new URLSearchParams({
    teamSiteId: args.teamSiteId,
    versionId: args.libraryContentVersionId,
    token,
  });
  if (args.context) {
    params.set("context", Buffer.from(args.context).toString("base64"));
  }
  if (args.prefillValues) {
    params.set("prefill", Buffer.from(JSON.stringify(args.prefillValues)).toString("base64"));
  }
  const url = `${FORM_APP_BASE}/form?${params}`;
  openWithDefaultApp(url);
  return {
    url,
    token,
    message: [
      `The form has been opened in the browser: ${url}`,
      `IMPORTANT: Call get_form_result with token="${token}" after the user submits the form.`,
      `get_form_result polls until the form app posts the result, then returns the generatedLivedocId and outputs.`,
    ].join("\n"),
  };
}

async function handleGetFormResult(args: { token: string }) {
  const res = await fetch(`${FORM_API_BASE}/api/result/${args.token}`);
  if (res.status === 404) {
    return { error: "Result not ready yet — the form may still be open or generation is in progress. Try again in a moment." };
  }
  if (!res.ok) {
    return { error: `Failed to retrieve result (HTTP ${res.status})` };
  }
  const data = await res.json() as { generatedLivedocId: string; outputs: Array<{ id: string; status: string; format: string; fileName: string }> };
  return {
    generatedLivedocId: data.generatedLivedocId,
    outputs: data.outputs,
    hint: "Generation complete. Call get_generation_download_url with generatedLivedocId and each output's id to get download links.",
  };
}

async function handleLogin(args: {
  username?: string;
  password?: string;
  tenant?: string;
  authServiceUri?: string;
  clientId?: string;
  clientSecret?: string;
}): Promise<{ ok: boolean; message: string } | { error: string; detail: unknown }> {
  const authUri  = args.authServiceUri ?? DEFAULT_AUTH_URI;
  const tenant   = args.tenant         ?? DEFAULT_AUTH_TENANT;
  const clientId     = args.clientId     ?? DEFAULT_CLIENT_ID;
  const clientSecret = args.clientSecret ?? DEFAULT_CLIENT_SECRET;
  const username = args.username ?? DEFAULT_USERNAME;
  const password = args.password ?? DEFAULT_PASSWORD;

  if (!authUri)  return { error: "authServiceUri is required (set AUTH_SERVICE_URI env var or pass authServiceUri).", detail: null };
  if (!tenant)   return { error: "tenant is required (set AUTH_TENANT env var or pass tenant).", detail: null };
  if (!clientId) return { error: "clientId is required (set AUTH_CLIENT_ID env var or pass clientId).", detail: null };
  if (!username) return { error: "username is required (set AUTH_USERNAME env var or pass username).", detail: null };
  if (!password) return { error: "password is required (set AUTH_PASSWORD env var or pass password).", detail: null };

  const tokenUrl = `${authUri}/tenants/${encodeURIComponent(tenant)}/connect/token`;
  const body = new URLSearchParams({
    grant_type:    "client_credentials",
    client_id:     clientId,
    client_secret: clientSecret,
    username,
    password,
    scope:         LOGIN_SCOPE,
  });

  const res = await fetch(tokenUrl, {
    method:  "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body:    body.toString(),
  });

  const text = await res.text();
  let data: Record<string, unknown>;
  try { data = JSON.parse(text) as Record<string, unknown>; } catch { return { error: `Auth server returned non-JSON (HTTP ${res.status})`, detail: text }; }

  if (!res.ok || !data.access_token) {
    return { error: `Login failed (HTTP ${res.status})`, detail: data };
  }

  currentToken = data.access_token as string;
  tokenIsManual = false;
  const expiresIn = data.expires_in as number | undefined;
  return { ok: true, message: `Token obtained successfully${expiresIn ? ` (expires in ${expiresIn}s)` : ""}. All tools are now authenticated.` };
}

// ── Server wiring ───────────────────────────────────────────────────────────

const server = new McpServer({ name: "seismic-livedoc", version: "1.0.0" });

// MCP App UI resource — served when Claude Desktop opens the App panel.
// frameDomains CSP is on the registration config (resources/list) so Claude Desktop
// applies it at connection time. The shell uses callServerTool only — no connectDomains needed.
registerAppResource(
  server,
  "LiveDoc Form",
  FORM_RESOURCE_URI,
  { description: "LiveDoc input form — React shell built by Vite." } as Parameters<typeof registerAppResource>[3],
  () => {
    const shellPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "dist", "views", "form-shell.html");
    const text = fs.readFileSync(shellPath, "utf-8");
    return { contents: [{ uri: FORM_RESOURCE_URI, mimeType: RESOURCE_MIME_TYPE, text }] };
  }
);

// ── Tool registrations ──────────────────────────────────────────────────────

server.registerTool(
  "search_livedoc_templates",
  {
    description:
      "Search for LiveDoc (Document Generator) templates in Seismic by name or keyword. Returns contentVersionId and teamSiteId needed for other tools. " +
      "ALWAYS call this FIRST whenever the user names or describes a template — do NOT ask for teamSiteId/libraryContentVersionId directly. " +
      "Only ask the user to disambiguate if this search returns zero or multiple plausible matches.",
    inputSchema: {
      searchText: z.string().optional().describe("Text to search across template title, description, and body."),
      page_size: z.number().optional().describe("Number of results to return (default 10, max 50)."),
    },
  },
  async (args) => {
    const result = await handleSearchTemplates(args as { searchText?: string; page_size?: number });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "search_livedoc_content",
  {
    description:
      "Search the Seismic library for content items by keyword. NOT part of the normal LiveDoc generation flow. " +
      "Only for standalone content discovery when the user explicitly asks to search for content.",
    inputSchema: {
      query: z.string().describe("Search text (title or keyword)."),
      contentType: z.enum(["ExternalSlides", "LiveSlide", "ExternalStaticSlides", "LiveDoc", "PDF"]).optional().describe("Filter by content type."),
      page_size: z.number().optional().describe("Number of results to return (default 10, max 50)."),
    },
  },
  async (args) => {
    const result = await handleSearchContent(args as { query: string; contentType?: string; page_size?: number });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

// get_livedoc_inputs — opens the MCP App panel with the form
registerAppTool(
  server,
  "get_livedoc_inputs",
  {
    description:
      "Retrieve the full input schema for a LiveDoc template and open the interactive input form in the App panel. " +
      "The form appears in the App panel automatically. The user fills it out and clicks Submit — " +
      "generation starts directly from the panel. " +
      "Once the user tells you generation is done (or the panel shows 'Generation complete'), " +
      "call get_panel_result with the formToken to get the generatedLivedocId and download URLs, " +
      "then call download_generation_output to save the file locally. " +
      "Do NOT build your own form, do NOT use AskUserQuestion.",
    inputSchema: {
      teamSiteId: z.string().describe("Team site identifier (UUID) that owns the template."),
      libraryContentVersionId: z.string().describe("Content version identifier (UUID) of the LiveDoc template."),
    },
    _meta: { ui: { resourceUri: FORM_RESOURCE_URI } },
  },
  async (args) => {
    const ir = await handleGetInputs(args as { teamSiteId: string; libraryContentVersionId: string });
    if ("error" in ir) {
      return { content: [{ type: "text" as const, text: JSON.stringify(ir, null, 2) }], isError: true };
    }

    const formToken = generateToken();
    const schema = buildFormSchema(ir, formToken) as Record<string, unknown> & {
      slideGroups: Array<{ thumbnailUrl: string; [k: string]: unknown }>;
    };

    // Slide group thumbnailUrls are Seismic signed CDN URLs — the App panel iframe
    // cannot load them cross-origin. Fetch server-side and inline as base64 data URLs.
    await Promise.all(
      (schema.slideGroups ?? []).map(async (g) => {
        if (!g.thumbnailUrl) return;
        try {
          const res = await fetch(g.thumbnailUrl, { headers: authHeaders() as Record<string, string> });
          if (!res.ok) return;
          const buf = Buffer.from(await res.arrayBuffer());
          const mime = res.headers.get("content-type") ?? "image/jpeg";
          g.thumbnailUrl = `data:${mime};base64,${buf.toString("base64")}`;
        } catch { /* leave original URL — panel will show broken img */ }
      })
    );

    pendingFormSchemas.set(formToken, schema);
    // Write to temp file so Process 2 (App panel) can read it
    fs.writeFileSync(
      path.join(os.tmpdir(), `mcp-livedoc-schema-${formToken}.json`),
      JSON.stringify(schema),
      "utf-8"
    );
    // Overwrite the "latest" pointer so the panel can detect a new generation request
    // even when the App panel was already open from a previous run.
    fs.writeFileSync(
      path.join(os.tmpdir(), `mcp-livedoc-latest-token.json`),
      JSON.stringify({ formToken }),
      "utf-8"
    );

    return {
      content: [{
        type: "text" as const,
        text: `Form opened in the App panel (formToken="${formToken}"). ` +
          `The user fills it out and clicks Submit in the panel — generation runs automatically. ` +
          `DO NOT call open_form_ui. DO NOT call submit_form. DO NOT call get_form_result. ` +
          `Just tell the user to fill the form. When they say it is done, call get_panel_result with formToken="${formToken}".`,
      }],
      structuredContent: { formToken },
    };
  }
);


server.registerTool(
  "log_debug_message",
  {
    description: "Internal: debug logging from the App panel. Do NOT call this yourself.",
    inputSchema: { msg: z.string() },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug.log"),
      `[${new Date().toISOString()}] APP_PANEL: ${(args as { msg: string }).msg}\n`);
    return { content: [{ type: "text" as const, text: "ok" }] };
  }
);

server.registerTool(
  "get_form_schema",
  {
    description: "Internal: returns the normalised form schema for the given token. Called by the App panel shell.",
    inputSchema: { token: z.string() },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    const { token } = args as { token: string };
    let schema = pendingFormSchemas.get(token);
    if (!schema) {
      // Process 2 (App panel) — read from temp file written by Process 1
      const schemaPath = path.join(os.tmpdir(), `mcp-livedoc-schema-${token}.json`);
      if (fs.existsSync(schemaPath)) {
        schema = JSON.parse(fs.readFileSync(schemaPath, "utf-8"));
        pendingFormSchemas.set(token, schema!);
      }
    }
    return {
      content: [{ type: "text" as const, text: schema ? "ok" : "schema-not-found" }],
      ...(schema ? { structuredContent: schema as Record<string, unknown> } : {}),
    };
  }
);

// Lets the already-open App panel detect a new generation request without remounting.
server.registerTool(
  "get_latest_token",
  {
    description: "Internal: returns the formToken for the most recent get_livedoc_inputs call. " +
      "The App panel polls this when idle/done so it can auto-reload when Claude triggers a new generation.",
    inputSchema: { currentToken: z.string().optional() },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    const { currentToken } = args as { currentToken?: string };
    const latestPath = path.join(os.tmpdir(), `mcp-livedoc-latest-token.json`);
    if (!fs.existsSync(latestPath)) {
      return { content: [{ type: "text" as const, text: "no-token" }], structuredContent: { formToken: null, isNew: false } };
    }
    const { formToken } = JSON.parse(fs.readFileSync(latestPath, "utf-8")) as { formToken: string };
    const isNew = !!formToken && formToken !== currentToken;
    return {
      content: [{ type: "text" as const, text: isNew ? `new-token:${formToken}` : "same" }],
      structuredContent: { formToken, isNew },
    };
  }
);

server.registerTool(
  "submit_form",
  {
    description: "Internal: called by the App panel after the user submits the form. Triggers LiveDoc generation.",
    inputSchema: {
      token: z.string(),
      payload: z.string().describe("JSON-serialised generation payload built by the form"),
    },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    const { token, payload } = args as { token: string; payload: string };
    let schema = pendingFormSchemas.get(token) as { teamSiteId: string; libraryContentVersionId: string; templateName?: string } | undefined;
    if (!schema) {
      const schemaPath = path.join(os.tmpdir(), `mcp-livedoc-schema-${token}.json`);
      if (fs.existsSync(schemaPath)) {
        schema = JSON.parse(fs.readFileSync(schemaPath, "utf-8"));
        pendingFormSchemas.set(token, schema!);
        try { fs.unlinkSync(schemaPath); } catch { /* ignore */ }
      }
    }
    if (!schema) {
      return { content: [{ type: "text" as const, text: "error: schema not found" }], structuredContent: { error: "Schema not found for token" } };
    }
    let parsed: Record<string, unknown>;
    try { parsed = JSON.parse(payload); } catch {
      return { content: [{ type: "text" as const, text: "error: invalid JSON" }], structuredContent: { error: "Invalid payload JSON" } };
    }
    const result = await handleSubmitGeneration({
      teamSiteId: schema.teamSiteId,
      libraryContentVersionId: schema.libraryContentVersionId,
      adHocInputs: (parsed.adHocInputs as Array<{ name: string; value: unknown }>) ?? [],
      outputs: (parsed.outputs as Array<{ format: string; fileName?: string }>) ?? [],
      variableListData: parsed.variableListData as never,
      manualSelectContentInput: parsed.manualSelectContentInput as never,
    });
    // Write to temp files so get_panel_result (chat) can pick up the generatedLivedocId
    const gid = (result as Record<string, unknown>).generatedLivedocId as string | undefined;
    if (gid) {
      pendingGenerations.set(gid, token);
      const resultData = JSON.stringify({ generatedLivedocId: gid, status: "Generating", downloadUrls: [], templateName: schema?.templateName ?? "" });
      fs.writeFileSync(path.join(os.tmpdir(), `mcp-livedoc-result-${token}.json`), resultData, "utf-8");
      // Reverse-lookup file survives server restarts
      fs.writeFileSync(path.join(os.tmpdir(), `mcp-livedoc-gid-${gid}.json`), JSON.stringify({ formToken: token }), "utf-8");
    }
    return {
      content: [{ type: "text" as const, text: JSON.stringify(result) }],
      structuredContent: result,
    };
  }
);

server.registerTool(
  "poll_generation",
  {
    description: "Internal: polls generation status. Called by App panel to track progress and get download URLs.",
    inputSchema: { generatedLivedocId: z.string() },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    const { generatedLivedocId } = args as { generatedLivedocId: string };
    const res = await apiFetch(`/v3/generatedLivedocs/${generatedLivedocId}`);
    const dbg = (msg: string) => fs.appendFileSync(
      path.join(os.tmpdir(), "mcp-livedoc-debug.log"),
      `[${new Date().toISOString()}] POLL: ${msg}\n`
    );
    if (res.status !== 200) {
      dbg(`HTTP ${res.status} for ${generatedLivedocId}`);
      return { content: [{ type: "text" as const, text: "error" }], structuredContent: { status: "Failed", error: `HTTP ${res.status}` } };
    }
    const body = res.body as Record<string, unknown>;
    const topStatus = statusName(body.status ?? body.Status);
    const outputs = ((body.outputs ?? body.Outputs) as Array<Record<string, unknown>>) ?? [];
    const outputStatuses = outputs.map(o => statusName(o.status ?? o.Status));
    const allDone = outputs.length > 0 && outputStatuses.every(s => s === "Completed" || s === "Failed");
    dbg(`id=${generatedLivedocId} topStatus=${topStatus} outputs=${JSON.stringify(outputStatuses)}`);

    // Consider done when ALL outputs have individually completed (top-level status can lag)
    const status = allDone
      ? (outputStatuses.some(s => s === "Failed") ? "Failed" : "Completed")
      : (topStatus === "Failed" ? "Failed" : "Generating");

    const downloadUrls: string[] = [];
    const downloads: Array<{ url: string; format: string; fileName: string }> = [];
    if (status === "Completed") {
      await Promise.all(outputs.map(async (o) => {
        const outputId = String(o.id ?? o.Id ?? "");
        const format  = String(o.format ?? o.Format ?? "pptx").toLowerCase();
        const rawName = String(o.fileName ?? o.FileName ?? `output`);
        const fileName = path.extname(rawName) ? rawName : `${rawName}.${format}`;
        if (!outputId) return;
        const dlResult = await handleGetDownloadUrl({ generatedLivedocId, outputId });
        dbg(`dl outputId=${outputId} result=${JSON.stringify(dlResult)}`);
        const dlBody = dlResult as Record<string, unknown>;
        const url = String(dlBody.url ?? dlBody.downloadUrl ?? dlBody.Url ?? dlBody.DownloadUrl ?? "");
        if (url) { downloadUrls.push(url); downloads.push({ url, format, fileName }); }
      }));

      // Update result file so get_panel_result (chat) can pick up the final URLs
      let formToken = pendingGenerations.get(generatedLivedocId);
      if (!formToken) {
        const lookupPath = path.join(os.tmpdir(), `mcp-livedoc-gid-${generatedLivedocId}.json`);
        if (fs.existsSync(lookupPath)) {
          try { formToken = (JSON.parse(fs.readFileSync(lookupPath, "utf-8")) as { formToken: string }).formToken; } catch { /* ignore */ }
        }
      }
      if (formToken) {
        // Preserve templateName written by submit_form
        let templateName = "";
        const existingResultPath = path.join(os.tmpdir(), `mcp-livedoc-result-${formToken}.json`);
        try { templateName = (JSON.parse(fs.readFileSync(existingResultPath, "utf-8")) as { templateName?: string }).templateName ?? ""; } catch { /* ignore */ }
        fs.writeFileSync(
          existingResultPath,
          JSON.stringify({ generatedLivedocId, status: "Completed", downloadUrls, downloads, templateName }),
          "utf-8"
        );
      }
    }
    return {
      content: [{ type: "text" as const, text: status }],
      structuredContent: {
        status, downloadUrls, downloads,
        // Include per-output detail so the panel can show a meaningful failure reason
        outputs: outputs.map(o => ({
          format: String(o.format ?? o.Format ?? ""),
          status: statusName(o.status ?? o.Status),
          errorMessage: String(o.errorMessage ?? o.ErrorMessage ?? o.error ?? o.Error ?? ""),
        })),
      },
    };
  }
);

server.registerTool(
  "download_output_file",
  {
    description: "Internal: downloads a generation output URL to the local Downloads folder and opens it. Called by the App panel download buttons.",
    inputSchema: {
      url: z.string(),
      fileName: z.string().optional().describe("Suggested filename with extension, e.g. 'MyDoc.pptx'."),
    },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    const { url, fileName } = args as { url: string; fileName?: string };
    try {
      const res = await fetch(url);
      if (!res.ok) {
        return { content: [{ type: "text" as const, text: `Download failed: HTTP ${res.status}` }], structuredContent: { error: `HTTP ${res.status}` } };
      }
      const buffer = await res.arrayBuffer();
      const name = fileName || "livedoc-output.pptx";
      const localPath = uniqueFilePath(getDownloadsDir(), name);
      fs.writeFileSync(localPath, Buffer.from(buffer));
      openWithDefaultApp(localPath);
      return {
        content: [{ type: "text" as const, text: `Saved to ${localPath}` }],
        structuredContent: { localPath, success: true },
      };
    } catch (e) {
      return { content: [{ type: "text" as const, text: `Error: ${e}` }], structuredContent: { error: String(e) } };
    }
  }
);

server.registerTool(
  "get_candidate_thumbnails",
  {
    description: "Internal: fetches top-level thumbnail URLs for a batch of content candidates. Called by the App panel after form schema loads.",
    inputSchema: {
      teamSiteId: z.string(),
      candidates: z.array(z.object({ contentId: z.string(), versionId: z.string() })),
    },
    _meta: { ui: { visibility: ["app"] } },
  },
  async (args) => {
    const { teamSiteId, candidates } = args as { teamSiteId: string; candidates: Array<{ contentId: string; versionId: string }> };
    const results = await Promise.all(
      candidates.map(async ({ contentId, versionId }) => {
        if (!contentId || !versionId) return { versionId, thumbnailUrl: "" };
        const res = await apiFetch("/v3/slides", {
          method: "POST",
          body: JSON.stringify({ teamSiteId, contentId, contentVersionId: versionId }),
        });
        if (res.status !== 200) return { versionId, thumbnailUrl: "" };
        const body = res.body as Record<string, unknown>;
        // imageUrl = top-level content thumbnail; contentThumbnailImageUrls[0] = first slide
        const thumbnailUrl = String(
          (body.contentThumbnailImageUrls as string[] | undefined)?.[0] ?? body.imageUrl ?? ""
        );
        return { versionId, thumbnailUrl };
      })
    );
    const thumbnailMap: Record<string, string> = {};
    for (const { versionId, thumbnailUrl } of results) {
      if (thumbnailUrl) thumbnailMap[versionId] = thumbnailUrl;
    }
    return {
      content: [{ type: "text" as const, text: `${Object.keys(thumbnailMap).length} thumbnails loaded` }],
      structuredContent: { thumbnailMap },
    };
  }
);

server.registerTool(
  "get_panel_result",
  {
    description:
      "Get the result of the LiveDoc generation triggered from the App panel. " +
      "Call this after the user says generation is done. Returns generatedLivedocId, status, downloads array, and templateName. " +
      "IMPORTANT — when status is 'Completed': " +
      "(1) Create an HTML artifact (type='text/html') showing a generation-complete card. " +
      "The card must include: a green check icon, 'Generation complete' heading, templateName, 'Completed' badge, " +
      "a DOWNLOADS section listing each file (icon by format, fileName, format label, a download arrow link to its url), " +
      "and a small 'Links expire …' note at the bottom. Keep the HTML concise (no external resources). " +
      "(2) Also call download_generation_output for each output to save the files locally. " +
      "If status is still 'Generating', tell the user to wait and offer to check again.",
    inputSchema: {
      formToken: z.string().describe("The formToken returned by get_livedoc_inputs."),
    },
  },
  async (args) => {
    const { formToken } = args as { formToken: string };
    const resultPath = path.join(os.tmpdir(), `mcp-livedoc-result-${formToken}.json`);
    if (!fs.existsSync(resultPath)) {
      return {
        content: [{ type: "text" as const, text: "No result yet — generation has not started or the form has not been submitted. Check the App panel." }],
      };
    }
    const data = JSON.parse(fs.readFileSync(resultPath, "utf-8")) as {
      generatedLivedocId: string;
      status: string;
      downloadUrls: string[];
      downloads?: Array<{ url: string; format: string; fileName: string }>;
      templateName?: string;
    };
    const text = data.status === "Completed"
      ? `Generation complete. templateName: "${data.templateName ?? ""}". generatedLivedocId: ${data.generatedLivedocId}. ` +
        `downloads: ${JSON.stringify(data.downloads ?? [])}. ` +
        `Create an HTML artifact showing the result card, then call download_generation_output for each output.`
      : `Generation status: ${data.status}. generatedLivedocId: ${data.generatedLivedocId}. Check back shortly.`;
    return {
      content: [{ type: "text" as const, text }],
      structuredContent: data,
    };
  }
);

server.registerTool(
  "submit_livedoc_generation",
  {
    description:
      "Submit a LiveDoc generation job. Provide ad hoc input values and at least one output format (PPTX, DOCX, PDF). Returns a generatedLivedocId to poll for status.",
    inputSchema: {
      teamSiteId: z.string().describe("Team site identifier (UUID)."),
      libraryContentVersionId: z.string().describe("Content version identifier (UUID) of the LiveDoc template."),
      adHocInputs: z.array(z.object({ name: z.string(), value: z.any() })).describe("Array of {name, value} pairs for ALL ad hoc inputs."),
      outputs: z.array(z.object({
        format: z.string(),
        name: z.string().optional(),
        fileName: z.string().describe("Filename with extension, e.g. \"Template.pdf\". Always set this."),
      })).describe("Output formats to generate."),
      variableListData: z.array(z.object({
        variableListName: z.string(),
        variableInputs: z.array(z.object({ name: z.string(), value: z.any() })),
      })).optional().describe("Variable list data from variableListDefinitions."),
      liveFormSellerTemplateId: z.string().optional(),
      regionalFormat: z.string().optional().describe("Regional format culture name, e.g. \"en-US\"."),
      manualSelectContentInput: z.object({
        manualSelectContentItems: z.array(z.object({
          id: z.string().describe("Stable slot identifier — copy verbatim from get_livedoc_inputs."),
          name: z.string().optional(),
          contentType: z.string().describe("One of \"Group\", \"Section\", \"LiveSlide\", \"ResourcePDF\", etc."),
          versionId: z.string().optional().describe("contentVersionId from the pasted form payload. Never populate yourself."),
          sourceBlobId: z.string().optional(),
          pageNumber: z.number().optional(),
          isInclude: z.boolean(),
          orderIndex: z.number().optional(),
        })),
      }).optional().describe("Content selection — pass ONLY what the pasted form payload contained."),
    },
  },
  async (args) => {
    const subResult = await handleSubmitGeneration(args as Parameters<typeof handleSubmitGeneration>[0]);
    const subBody = subResult as Record<string, unknown>;
    if (subBody.error) {
      return { content: [{ type: "text" as const, text: JSON.stringify(subBody, null, 2) }], isError: true };
    }
    const gid = String(subBody.generatedLivedocId ?? "");
    return {
      content: [{
        type: "text" as const,
        text: [
          `Generation started. generatedLivedocId: ${gid}`,
          `NEXT: call get_generation_status with generatedLivedocId="${gid}". Keep calling every few seconds until allDone=true.`,
        ].join("\n"),
      }],
    };
  }
);

server.registerTool(
  "get_generation_status",
  {
    description: "Check the status of a LiveDoc generation job. Poll until all outputs reach 'Completed' or 'Failed'.",
    inputSchema: {
      generatedLivedocId: z.string().describe("The generatedLivedocId returned by submit_livedoc_generation."),
    },
  },
  async (args) => {
    const statusResult = await handleGetStatus(args as { generatedLivedocId: string });
    const st = statusResult as Record<string, unknown>;
    if (st.error) {
      return { content: [{ type: "text" as const, text: JSON.stringify(st, null, 2) }], isError: true };
    }
    const outputs = (st.outputs as Array<Record<string, unknown>>) ?? [];
    const allDone = st.allDone as boolean;
    const nextStep = allDone
      ? `All done. NEXT: call download_generation_output for each completed output:\n${outputs.filter(o => o.status === "Completed").map(o => `  outputId="${o.id}" (${o.format} — ${o.fileName})`).join("\n")}`
      : `Still generating. NEXT: call get_generation_status again with generatedLivedocId="${st.generatedLivedocId}" in a few seconds.`;
    return {
      content: [{
        type: "text" as const,
        text: [JSON.stringify(st, null, 2), nextStep].join("\n\n"),
      }],
    };
  }
);

server.registerTool(
  "open_form_ui",
  {
    description:
      "DEPRECATED — DO NOT call this after get_livedoc_inputs. The form is now embedded in the App panel. " +
      "get_livedoc_inputs already opens the panel form automatically. Calling this tool will open a redundant browser window. " +
      "This tool is kept only as a last-resort fallback when the App panel is unavailable.",
    inputSchema: {
      teamSiteId: z.string().describe("Team site identifier (UUID)."),
      libraryContentVersionId: z.string().describe("Content version identifier (UUID) of the LiveDoc template."),
      context: z.string().optional().describe("The user's original generation request (natural language)."),
      prefillValues: z.any().optional().describe("Optional AI-suggested default values to pre-populate the form."),
    },
  },
  async (args) => {
    const result = await handleOpenFormUi(args as { teamSiteId: string; libraryContentVersionId: string; context?: string; prefillValues?: unknown });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "get_form_result",
  {
    description: "Retrieve the generation result posted back by the Form UI after the user completed and closed the form.",
    inputSchema: {
      token: z.string().describe("The token returned by open_form_ui."),
    },
  },
  async (args) => {
    const result = await handleGetFormResult(args as { token: string });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "login",
  {
    description:
      "Obtain a Seismic bearer token using username and password (OAuth 2.0 Resource Owner Password Credentials). Automatically sets the token for all subsequent tool calls.",
    inputSchema: {
      username: z.string().optional().describe(`Seismic username. Defaults to AUTH_USERNAME env var${DEFAULT_USERNAME ? " (pre-configured)" : ""}.`),
      password: z.string().optional().describe(`Seismic password. Defaults to AUTH_PASSWORD env var.`),
      tenant: z.string().optional().describe(`Tenant slug, e.g. "qa01eastasia01". Defaults to AUTH_TENANT env var (${DEFAULT_AUTH_TENANT || "not set"}).`),
      authServiceUri: z.string().optional().describe("Auth service base URL. Defaults to AUTH_SERVICE_URI env var."),
      clientId: z.string().optional().describe("OAuth client ID."),
      clientSecret: z.string().optional().describe("OAuth client secret."),
    },
  },
  async (args) => {
    const result = await handleLogin(args as Parameters<typeof handleLogin>[0]);
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "set_token",
  {
    description: "Update the Seismic API bearer token used by all other tools. Call this whenever a tool returns HTTP 401 before retrying.",
    inputSchema: {
      token: z.string().describe("The new bearer token (without the 'Bearer ' prefix)."),
    },
  },
  async (args) => {
    currentToken = (args as { token: string }).token;
    tokenIsManual = true;
    return { content: [{ type: "text" as const, text: "Token updated. This token will not be auto-replaced by the credential-flow login on a 401." }] };
  }
);

server.registerTool(
  "get_generation_download_url",
  {
    description:
      "Get the download URL for a completed LiveDoc output. Returns a JSON payload with the URL. Use outputId from get_generation_status, or a format alias like 'pptx', 'docx', 'pdf'.",
    inputSchema: {
      generatedLivedocId: z.string().describe("The generatedLivedocId."),
      outputId: z.string().describe("The output id or a format alias: 'pptx', 'docx', 'pdf', 'gslides', 'gdoc'."),
    },
  },
  async (args) => {
    const dlResult = await handleGetDownloadUrl(args as { generatedLivedocId: string; outputId: string });
    if (dlResult && typeof dlResult === "object" && "error" in (dlResult as object)) {
      return { content: [{ type: "text" as const, text: JSON.stringify(dlResult, null, 2) }], isError: true };
    }
    const dlBody = dlResult as Record<string, unknown>;
    const dlUrl = String(dlBody.url ?? dlBody.downloadUrl ?? dlBody.Url ?? dlBody.DownloadUrl ?? "");
    const dlFile = String(dlBody.fileName ?? dlBody.FileName ?? dlBody.name ?? dlBody.Name ?? "download");
    if (!dlUrl) {
      return { content: [{ type: "text" as const, text: JSON.stringify(dlResult, null, 2) }] };
    }
    return {
      content: [{
        type: "text" as const,
        text: `✅ **${dlFile}** is ready.\n\nDownload link: [${dlFile}](${dlUrl})\n\n(Reproduce the markdown link above verbatim in your reply so the user can click it.)`,
      }],
    };
  }
);

server.registerTool(
  "download_generation_output",
  {
    description:
      "Download a completed LiveDoc output to the Working folder and open it with the system default app.",
    inputSchema: {
      generatedLivedocId: z.string().describe("The generatedLivedocId."),
      outputId: z.string().describe("The output id or a format alias: 'pptx', 'docx', 'pdf', 'gslides', 'gdoc'."),
      autoOpen: z.boolean().optional().describe("Whether to automatically open the file. Default true."),
    },
  },
  async (args) => {
    const result = await handleDownloadGenerationOutput(args as { generatedLivedocId: string; outputId: string; autoOpen?: boolean });
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.registerTool(
  "debug_environment",
  {
    description:
      "Diagnostic tool: dumps process working directory, matched environment variable names/values, and client capabilities.",
    inputSchema: {},
  },
  async () => {
    const keywords = /claude|anthropic|output|sandbox|session|cowork|agent|workspace/i;
    const looksLikePath = (v: string) => /^[a-zA-Z]:[\\/]|^\//.test(v) && /[\\/]/.test(v);
    const matchedVarNames = Object.keys(process.env).filter((k) => keywords.test(k));
    const pathLikeValues = Object.fromEntries(
      matchedVarNames
        .map((k) => [k, process.env[k] ?? ""])
        .filter(([, v]) => looksLikePath(v as string))
    );
    const clientCaps = server.server.getClientCapabilities();
    const result = {
      cwd: process.cwd(),
      matchedEnvVarNames: matchedVarNames,
      pathLikeEnvVarValues: pathLikeValues,
      clientCapabilitiesRawJSON: JSON.stringify(clientCaps, null, 2),
      elicitationSupported: !!(clientCaps?.elicitation),
      mcpAppUiSupported: !!(clientCaps && JSON.stringify(clientCaps).includes("mcp-app")),
    };
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

// ── PPTX Auto-Tagging tools (PoC) ────────────────────────────────────────────

server.registerTool(
  "pptx_extract_shapes",
  {
    description:
      "Extract all shapes from a PPTX file and return a structured list. Requires the local PoC server (cd poc-auto-tagging && npm start).",
    inputSchema: {
      pptxBase64: z.string().describe("Base64-encoded PPTX file content."),
      slideIndex: z.number().optional().describe("Optional: only extract shapes from this slide (0-based)."),
    },
  },
  async (args) => {
    const a = args as { pptxBase64: string; slideIndex?: number };
    const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
    const r = await fetch(`${pocBase}/api/pptx/extract`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pptxBase64: a.pptxBase64, ...(a.slideIndex !== undefined ? { slideIndex: a.slideIndex } : {}) }),
    });
    if (!r.ok) throw new Error(`pptx_extract_shapes HTTP ${r.status}: ${await r.text()}`);
    return { content: [{ type: "text" as const, text: JSON.stringify(await r.json(), null, 2) }] };
  }
);

server.registerTool(
  "pptx_auto_tag",
  {
    description:
      "Send a PPTX to the AI (local LLM) for automatic analysis. Requires the local PoC server (cd poc-auto-tagging && npm start).",
    inputSchema: {
      pptxBase64: z.string().describe("Base64-encoded PPTX file content."),
      schema: z.record(z.string(), z.string()).optional().describe("Optional datasource schema as a JSON object."),
    },
  },
  async (args) => {
    const a = args as { pptxBase64: string; schema?: Record<string, string> };
    const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
    const r = await fetch(`${pocBase}/api/pptx/auto-tag`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pptxBase64: a.pptxBase64, schema: a.schema ?? {} }),
    });
    if (!r.ok) throw new Error(`pptx_auto_tag HTTP ${r.status}: ${await r.text()}`);
    return { content: [{ type: "text" as const, text: JSON.stringify(await r.json(), null, 2) }] };
  }
);

server.registerTool(
  "pptx_mark_shapes",
  {
    description:
      "Apply dynamic element markings to a PPTX file. Requires the local PoC server (cd poc-auto-tagging && npm start).",
    inputSchema: {
      pptxBase64: z.string().describe("Base64-encoded PPTX file content."),
      marks: z.array(z.object({
        slideIndex: z.number().describe("0-based slide index."),
        shapeId: z.number().describe("Numeric shape ID from pptx_extract_shapes."),
        varName: z.string().describe("camelCase variable name, e.g. 'companyName'."),
        varType: z.enum(["text", "image", "table", "chart", "number", "date"]),
        description: z.string().optional(),
      })).describe("List of shapes to mark as dynamic."),
    },
  },
  async (args) => {
    const a = args as { pptxBase64: string; marks: unknown[] };
    const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
    const r = await fetch(`${pocBase}/api/pptx/mark`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pptxBase64: a.pptxBase64, marks: a.marks }),
    });
    if (!r.ok) throw new Error(`pptx_mark_shapes HTTP ${r.status}: ${await r.text()}`);
    const result = await r.json() as { bindings: unknown[]; markedAt: string };
    return {
      content: [{
        type: "text" as const,
        text: JSON.stringify({ markedAt: result.markedAt, bindingCount: result.bindings.length, bindings: result.bindings }, null, 2) +
          "\n\n(markedPptxBase64 available in full result — use pptx_mark_shapes result.pptxBase64 to save the file)",
      }],
    };
  }
);

server.registerTool(
  "pptx_get_manifest",
  {
    description:
      "Extract the current binding manifest from a PPTX file. Requires the local PoC server (cd poc-auto-tagging && npm start).",
    inputSchema: {
      pptxBase64: z.string().describe("Base64-encoded PPTX file content."),
    },
  },
  async (args) => {
    const a = args as { pptxBase64: string };
    const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
    const r = await fetch(`${pocBase}/api/pptx/manifest?pptxBase64=${encodeURIComponent(a.pptxBase64)}`);
    if (!r.ok) throw new Error(`pptx_get_manifest HTTP ${r.status}: ${await r.text()}`);
    return { content: [{ type: "text" as const, text: JSON.stringify(await r.json(), null, 2) }] };
  }
);

process.on("SIGTERM", () => process.exit(0));
process.on("SIGINT", () => process.exit(0));
process.on("uncaughtException", (err) => {
  fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-crash.log"), `[${new Date().toISOString()}] uncaughtException: ${err.stack ?? err.message}\n`);
  process.exit(1);
});
process.on("unhandledRejection", (reason) => {
  fs.appendFileSync(path.join(os.tmpdir(), "mcp-livedoc-crash.log"), `[${new Date().toISOString()}] unhandledRejection: ${reason}\n`);
  process.exit(1);
});

// Auto-login on startup when credentials are available via env vars
if (!currentToken && DEFAULT_USERNAME && DEFAULT_PASSWORD) {
  await autoLogin();
}

const debugLog = path.join(os.tmpdir(), "mcp-livedoc-debug.log");
server.server.oninitialized = () => {
  const caps = server.server.getClientCapabilities();
  const uiCap = getUiCapability(caps as Parameters<typeof getUiCapability>[0]);
  fs.appendFileSync(debugLog, `[${new Date().toISOString()}] oninitialized: extensions=${JSON.stringify(caps?.extensions)}, uiCap=${JSON.stringify(uiCap)}\n`);
};

const transport = new StdioServerTransport();
await server.connect(transport);
