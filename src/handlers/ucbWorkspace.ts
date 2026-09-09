import { apiFetch } from "../api/client.js";
import { getToken } from "../auth/state.js";
import { jwtTenantFqdn } from "../auth/jwt.js";
import { INTERNAL_BASE_URL } from "../config.js";

// LDS's own APIs never return a browsable URL for a committed file, so this is built client-side
// from fileId + the JWT's tenant_fqdn. `viewType` is NOT format-dependent (there is no per-format
// value — confirmed against workspace-service's ItemLocation enum, which only has SharedWithMe,
// MyFiles, TeamFolder, DraftPresentations: c:\project_new\workspace-service\src\Seismic.Workspace.Model\Location\ItemLocation.cs).
// It reflects the item's actual location context, determined server-side by
// GET /api/workspace/v2/tenants/{tenantId}/items/{itemId}/location — an internal workspace-service
// endpoint LDS does not currently proxy. "DraftPresentations" is the one value confirmed to work
// end-to-end against a live QA tenant; treat it as a default, not a guaranteed-correct value for
// every destination folder.
function buildWorkspaceUrl(tenantFqdn: string, fileId: string): string {
  return `https://${tenantFqdn}/apps/workspace/doc/${fileId}//grid/title?viewType=DraftPresentations`;
}

// Keyed by generationId. Populated by submitUcbWorkspaceGeneration, consumed by
// getUcbWorkspaceGenerationStatus once the generation reaches "Ready" so the Workspace
// commit call can run without the model having to carry all these ids across turns.
// In-memory only: if the server restarts mid-flow the pending commit is lost and the
// generation must be resubmitted (generations are short-lived, so this is an acceptable trade-off).
interface PendingCommit {
  spaceId: string;
  fileId: string;
  fileVersionId: string;
  instanceId: string;
  stageId: string;
  stageRecordId: string;
  committed: boolean;
}

const pendingCommits = new Map<string, PendingCommit>();

// GetWorkspaceDestinationSpaces/Roots/FolderItems live in the "Document Generator (Internal)"
// API resource, not the main LiveDoc one — must use INTERNAL_BASE_URL. See config.ts.
export async function listWorkspaceSpaces() {
  const result = await apiFetch("/v3/workspace/destinations/spaces", {}, true, INTERNAL_BASE_URL);
  if (result.status !== 200) {
    return { error: `Listing Workspace spaces failed (HTTP ${result.status})`, detail: result.body };
  }
  return result.body;
}

export async function listWorkspaceFolders(args: {
  spaceId: string;
  folderId?: string;
  offset?: number;
  limit?: number;
}) {
  const offset = args.offset ?? 0;
  const limit = args.limit ?? 100;
  const path = args.folderId
    ? `/v3/workspace/destinations/spaces/${encodeURIComponent(args.spaceId)}/folders/${encodeURIComponent(args.folderId)}/items?offset=${offset}&limit=${limit}`
    : `/v3/workspace/destinations/spaces/${encodeURIComponent(args.spaceId)}/roots`;
  const result = await apiFetch(path, {}, true, INTERNAL_BASE_URL);
  if (result.status !== 200) {
    return { error: `Listing Workspace folder contents failed (HTTP ${result.status})`, detail: result.body };
  }
  return result.body;
}

export async function submitUcbWorkspaceGeneration(args: {
  teamSiteId: string;
  libraryContentVersionId: string;
  adHocInputs: Array<{ name: string; value: unknown }>;
  outputs: Array<{ format: string; name?: string; fileName?: string }>;
  variableListData?: Array<{
    variableListName: string;
    variableInputs: Array<{ name: string; value: unknown }>;
  }>;
  regionalFormat?: string;
  workspace: { spaceId: string; folderId: string; name: string; format: string };
  origin: { profileId: string; profileVersionId: string; contentLocation: string };
}) {
  if (args.outputs.length !== 1) {
    return {
      error: "Exactly one output is required for a UCB Workspace generation.",
      detail: `Got ${args.outputs.length} outputs.`,
    };
  }

  const generationInput: Record<string, unknown> = {
    adHocInputs: args.adHocInputs,
    outputs: args.outputs,
  };
  if (args.variableListData) generationInput.variableListData = args.variableListData;
  if (args.regionalFormat) generationInput.regionalFormat = args.regionalFormat;

  const reqBody = {
    generationInput,
    workspace: {
      spaceId: args.workspace.spaceId,
      folderId: args.workspace.folderId,
      name: args.workspace.name,
      format: args.workspace.format,
    },
    origin: {
      profileId: args.origin.profileId,
      profileVersionId: args.origin.profileVersionId,
      contentLocation: args.origin.contentLocation,
    },
  };

  const result = await apiFetch(
    `/v3/teamsites/${args.teamSiteId}/livedocVersions/${args.libraryContentVersionId}/ucb-workspace-generations`,
    { method: "POST", body: JSON.stringify(reqBody) }
  );
  if (result.status !== 201 && result.status !== 200) {
    return { error: `UCB Workspace generation submission failed (HTTP ${result.status})`, detail: result.body };
  }

  const body = result.body as Record<string, unknown>;
  const generationId = String(body.id ?? body.Id ?? "");
  const lifecycle = (body.lifecycle ?? body.Lifecycle ?? {}) as Record<string, unknown>;
  const workspace = (body.workspace ?? body.Workspace ?? {}) as Record<string, unknown>;

  const instanceId = String(lifecycle.instanceId ?? lifecycle.InstanceId ?? "");
  const stageId = String(lifecycle.stageId ?? lifecycle.StageId ?? "");
  const stageRecordId = String(lifecycle.stageRecordId ?? lifecycle.StageRecordId ?? "");
  const fileId = String(workspace.fileId ?? workspace.FileId ?? "");
  const fileVersionId = String(workspace.fileVersionId ?? workspace.FileVersionId ?? "");

  if (generationId && instanceId && stageId && stageRecordId && fileId && fileVersionId) {
    pendingCommits.set(generationId, {
      spaceId: args.workspace.spaceId,
      fileId,
      fileVersionId,
      instanceId,
      stageId,
      stageRecordId,
      committed: false,
    });
  } else {
    return {
      generationId,
      workspaceFileName: args.workspace.name,
      rawBody: body,
      warning: "Could not extract lifecycle/workspace ids (instanceId/stageId/stageRecordId/fileId/fileVersionId) " +
        "from the submission response. Auto-commit to Workspace will not work for this generationId — " +
        "get_ucb_workspace_generation_status will report the commit context as lost even without a server restart.",
    };
  }

  return {
    generationId,
    workspaceFileName: args.workspace.name,
    rawBody: body,
    message: "UCB Workspace generation submitted. Call get_ucb_workspace_generation_status to poll for completion.",
  };
}

async function commitToWorkspace(generationId: string, pending: PendingCommit) {
  const result = await apiFetch(
    `/v3/workspace/spaces/${encodeURIComponent(pending.spaceId)}/files/${encodeURIComponent(pending.fileId)}/versions/${encodeURIComponent(pending.fileVersionId)}/livedoc/instance`,
    {
      method: "POST",
      body: JSON.stringify({
        id: pending.instanceId,
        stage: { id: pending.stageId, recordId: pending.stageRecordId },
        useCustomName: false,
      }),
    }
  );
  if (result.status !== 200 && result.status !== 202 && result.status !== 204) {
    return { error: `Committing the generated file to Workspace failed (HTTP ${result.status})`, detail: result.body };
  }
  pending.committed = true;
  return { committed: true };
}

// Polling budget for a single tool call: previously each call to this tool did exactly one
// status check, so a caller with a limited number of tool-call rounds (e.g. livedoc-agent's
// MAX_TOOL_ROUNDS) could exhaust its whole budget polling a slow-to-finish generation and
// never see status "Ready" (and therefore never see workspaceUrl) at all. Looping internally
// here collapses "poll every couple seconds until done" into one tool call for the common case,
// so the caller only needs to re-invoke this tool if the generation is unusually slow.
const STATUS_POLL_BUDGET_MS = 25_000;
const STATUS_POLL_INTERVAL_MS = 2_000;

async function fetchStatusOnce(
  generationId: string
): Promise<{ ok: true; body: Record<string, unknown> } | { ok: false; error: string; detail?: unknown }> {
  const result = await apiFetch(`/v3/ucb-workspace-generations/${generationId}/status`);
  if (result.status !== 200) {
    return { ok: false, error: `Status check failed (HTTP ${result.status})`, detail: result.body };
  }
  return { ok: true, body: result.body as Record<string, unknown> };
}

export async function getUcbWorkspaceGenerationStatus(args: { generationId: string }) {
  const deadline = Date.now() + STATUS_POLL_BUDGET_MS;
  let raw: Record<string, unknown>;

  while (true) {
    const fetched = await fetchStatusOnce(args.generationId);
    if (!fetched.ok) return { error: fetched.error, detail: fetched.detail };
    raw = fetched.body;
    const status = String(raw.status ?? raw.Status ?? "");
    // Terminal states resolve immediately regardless of budget; only a non-terminal status
    // ("Queued"/"Generating"/...) is worth waiting out.
    if (status === "Ready" || Boolean(raw.isCompleted ?? raw.IsCompleted) || Date.now() >= deadline) {
      break;
    }
    await new Promise(r => setTimeout(r, STATUS_POLL_INTERVAL_MS));
  }

  const status = String(raw.status ?? raw.Status ?? "");
  const isCompleted = Boolean(raw.isCompleted ?? raw.IsCompleted);

  const response: Record<string, unknown> = {
    generationId: String(raw.id ?? raw.Id ?? args.generationId),
    status,
    isCompleted,
    formRecordId: raw.formRecordId ?? raw.FormRecordId ?? null,
    workspaceCommitted: false,
  };

  if (isCompleted && status !== "Ready") {
    const errorMessage = String(raw.errorMessage ?? raw.ErrorMessage ?? "").trim();
    return {
      ...response,
      error: `UCB Workspace generation ended with status "${status}"${errorMessage ? `: ${errorMessage}` : ""}`,
    };
  }
  if (status !== "Ready") {
    return {
      ...response,
      message: `Generation is still "${status}" after ${STATUS_POLL_BUDGET_MS / 1000}s of polling. ` +
        "Call get_ucb_workspace_generation_status again to keep waiting — do not report a workspaceUrl yet.",
    };
  }

  const pending = pendingCommits.get(args.generationId);
  if (!pending) {
    return {
      ...response,
      error: "Generation is Ready, but the Workspace commit context for this generationId was lost " +
        "(likely an MCP server restart mid-flow). The generation must be resubmitted via submit_ucb_workspace_generation.",
    };
  }
  if (pending.committed) {
    response.workspaceCommitted = true;
    response.workspaceUrl = buildWorkspaceUrlForPending(pending);
    return response;
  }

  const commitResult = await commitToWorkspace(args.generationId, pending);
  if ("error" in commitResult) {
    return {
      ...response,
      workspaceCommitted: false,
      commitError: commitResult.error,
      commitErrorDetail: commitResult.detail,
    };
  }
  response.workspaceCommitted = true;
  response.workspaceUrl = buildWorkspaceUrlForPending(pending);
  return response;
}

function buildWorkspaceUrlForPending(pending: PendingCommit): string | null {
  const tenantFqdn = jwtTenantFqdn(getToken());
  if (!tenantFqdn) return null;
  return buildWorkspaceUrl(tenantFqdn, pending.fileId);
}
