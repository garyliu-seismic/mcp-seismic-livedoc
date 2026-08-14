import { apiFetch, gf } from "../api/client.js";
import { dbg } from "../utils/debug.js";

export const CANDIDATE_PAGE_SIZE = 10;

export async function handleSearchTemplates(args: {
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
    totalCount: data.totalCount ?? 0,
    results: (data.documents ?? []).map((d) => ({
      title: d.title,
      format: d.format,
      contentVersionId: d.contentVersionId,
      teamSiteId: d.teamsite,
      modifiedDate: d.modifiedDate,
      description: d.description,
    })),
  };
}

export async function handleSearchContent(args: {
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
    totalCount: data.totalCount ?? 0,
    results: (data.documents ?? []).map((d) => ({
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
export function needsContentResolution(contentType: string): boolean {
  return contentType !== "" && contentType !== "Group" && contentType !== "Section";
}

// C# bool property names here don't follow simple camelCase (AllowPDF, IncludeStandardPPTX),
// so check several literal casings rather than relying on gf()'s single-fallback capitalization.
export function boolField(item: Record<string, unknown>, ...keys: string[]): boolean {
  for (const k of keys) {
    if (typeof item[k] === "boolean") return item[k] as boolean;
  }
  return false;
}

// Resolves real content candidates for one manualSelectContentItem. Prefers the item's own
// filter/format flags (the template author's actual search criteria, e.g. Filter: [{propertyName:
// "ContentName", operator: "CT", value: "sp3"}]) over a generic name-based guess — those flags
// are what get_livedoc_inputs actually returns on ExternalSlideContent items.
export async function resolveManualSelectCandidates(item: Record<string, unknown>): Promise<{ candidates: Array<Record<string, unknown>>; totalCount: number }> {
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
