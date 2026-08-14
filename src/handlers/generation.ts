import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { apiFetch } from "../api/client.js";
import { authHeaders } from "../auth/headers.js";
import { BASE_URL } from "../config.js";
import { getDownloadsDir, openWithDefaultApp, uniqueFilePath } from "../utils/os-utils.js";
import { needsContentResolution } from "./content.js";

// Matches LiveDocGenStatusResp in app-livedoc-service (PublicAPIV3Controller.ResultStatus.cs):
// Queued=0, Generating=1, Completed=2, Failed=3. The API returns this as a number, not a string,
// so callers must map it before comparing against status names.
export const STATUS_NAMES = ["Queued", "Generating", "Completed", "Failed"];

export function statusName(raw: unknown): string {
  if (typeof raw === "number" && STATUS_NAMES[raw] !== undefined) {
    return STATUS_NAMES[raw];
  }
  if (typeof raw === "string" && STATUS_NAMES.includes(raw)) {
    return raw;
  }
  return String(raw);
}

export async function handleSubmitGeneration(args: {
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

  // Debug dump — readable at %TEMP%\mcp-livedoc-debug-submit.json after each Submit
  try { fs.writeFileSync(path.join(os.tmpdir(), "mcp-livedoc-debug-submit.json"), JSON.stringify(reqBody, null, 2)); } catch { /* ignore */ }

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

export async function handleGetStatus(args: { generatedLivedocId: string }) {
  const result = await apiFetch(`/v3/generatedLivedocs/${args.generatedLivedocId}`);
  if (result.status !== 200) {
    return { error: `Status check failed (HTTP ${result.status})`, detail: result.body };
  }
  const raw = result.body as Record<string, unknown>;
  const id = (raw.id ?? raw.Id ?? raw.generatedLivedocId ?? raw.GeneratedLivedocId) as string;
  const rawOutputs = ((raw.outputs ?? raw.Outputs ?? []) as Array<Record<string, unknown>>)
    .filter(o => String(o.format ?? o.Format ?? "").toLowerCase() !== "thumbnail");
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

export async function handleGetDownloadUrl(args: {
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

export async function handleDownloadGenerationOutput(args: {
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
