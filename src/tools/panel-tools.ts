import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { dbg, DEBUG_LOG } from "../utils/debug.js";
import { getToken, setToken, isTokenManual, setTokenManual } from "../auth/state.js";
import { jwtExpiresAt } from "../auth/jwt.js";
import { handleLogin } from "../auth/auto-login.js";
import { apiFetch } from "../api/client.js";
import { handleGetInputs, buildFormSchema } from "../handlers/inputs.js";
import { statusName, handleSubmitGeneration, handleGetDownloadUrl } from "../handlers/generation.js";
import { generateToken, getDownloadsDir, openWithDefaultApp, uniqueFilePath } from "../utils/os-utils.js";
import {
  pendingFormSchemas, pendingGenerations,
  readSchema, writeSchema, deleteSchemaFile,
  writeLatestToken, writeResult, readResult, writeGid, readGid,
} from "../ipc/temp-file.js";
import { DEFAULT_AUTH_TENANT, DEFAULT_USERNAME, DEFAULT_PASSWORD, FORM_RESOURCE_URI } from "../config.js";

export function registerPanelTools(server: McpServer): void {
  // open_livedoc_panel — opens the App panel so the user can sign in or check status
  registerAppTool(
    server,
    "open_livedoc_panel",
    {
      description:
        "Open the LiveDoc App panel for sign-in. ONLY call this when a tool explicitly returns an HTTP 401 error, or when the user explicitly asks to log in. " +
        "Do NOT call this proactively before attempting any tool — always try the actual tool first and react to failures. " +
        "IMPORTANT: After calling this tool, you MUST stop and tell the user to sign in via the panel, then WAIT. " +
        "Do NOT call any other tools until the user sends a follow-up message confirming they have signed in. " +
        "Do NOT ask the user for credentials — the panel has its own sign-in form.",
      inputSchema: {},
      _meta: { ui: { resourceUri: FORM_RESOURCE_URI } },
    },
    async () => {
      const tok = getToken();
      const exp = tok ? jwtExpiresAt(tok) : null;
      const isAuthenticated = !!tok && (exp === null || Date.now() < exp - 60_000);
      return {
        content: [{ type: "text" as const, text: isAuthenticated
          ? "Panel opened. The user is already signed in — proceed with their request."
          : "Panel opened showing the sign-in form. STOP HERE. Tell the user to fill in their credentials in the panel and click Sign in. Do not call any other tool until the user confirms they have signed in.",
        }],
        structuredContent: { action: isAuthenticated ? "ready" : "show_login" },
      };
    }
  );

  server.registerTool(
    "check_auth",
    {
      description: "Internal: check auth status. Called by the App panel only.",
      inputSchema: {},
      _meta: { ui: { visibility: ["app"] } },
    },
    async () => {
      const tok = getToken();
      const exp = tok ? jwtExpiresAt(tok) : null;
      const isAuthenticated = !!tok && (exp === null || Date.now() < exp - 60_000);
      return {
        content: [{ type: "text" as const, text: isAuthenticated ? "Authenticated — token is valid." : "Not authenticated — token is missing or expired. Call open_livedoc_panel and wait for the user to sign in." }],
        structuredContent: { isAuthenticated, expiresAt: exp ?? null },
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
    "get_auth_status",
    {
      description: "Internal: returns whether the server currently holds a valid, non-expired bearer token. Called by the App panel on startup.",
      inputSchema: {},
      _meta: { ui: { visibility: ["app"] } },
    },
    async () => {
      const tok = getToken();
      if (!tok) {
        return { content: [{ type: "text" as const, text: "not-authenticated" }], structuredContent: { isAuthenticated: false } };
      }
      const exp = jwtExpiresAt(tok);
      if (exp !== null && Date.now() >= exp - 60_000) {
        return { content: [{ type: "text" as const, text: "token-expired" }], structuredContent: { isAuthenticated: false } };
      }
      return {
        content: [{ type: "text" as const, text: "authenticated" }],
        structuredContent: { isAuthenticated: true, expiresAt: exp ?? null },
      };
    }
  );

  server.registerTool(
    "get_auth_config",
    {
      description: "Internal: returns pre-configured auth values to pre-populate the login form. Called by the App panel on login screen load.",
      inputSchema: {},
      _meta: { ui: { visibility: ["app"] } },
    },
    async () => {
      return {
        content: [{ type: "text" as const, text: "ok" }],
        structuredContent: {
          tenant:      DEFAULT_AUTH_TENANT || null,
          hasUsername: !!DEFAULT_USERNAME,
          hasPassword: !!DEFAULT_PASSWORD,
        },
      };
    }
  );

  server.registerTool(
    "panel_login",
    {
      description: "Internal: authenticates with Seismic using tenant, username, and password. Called by the App panel login form.",
      inputSchema: {
        tenant:   z.string(),
        username: z.string(),
        password: z.string(),
      },
      _meta: { ui: { visibility: ["app"] } },
    },
    async (args) => {
      const { tenant, username, password } = args as { tenant: string; username: string; password: string };
      const result = await handleLogin({ tenant, username, password });
      if ("error" in result) {
        return {
          content: [{ type: "text" as const, text: result.error }],
          structuredContent: { ok: false, error: result.error },
          isError: true,
        };
      }
      return {
        content: [{ type: "text" as const, text: result.message }],
        structuredContent: { ok: true },
      };
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
      const schema = readSchema(token);
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
      const { currentToken: panelToken } = args as { currentToken?: string };
      const data = (() => {
        try {
          const latestPath = path.join(os.tmpdir(), `mcp-livedoc-latest-token.json`);
          if (!fs.existsSync(latestPath)) return null;
          return JSON.parse(fs.readFileSync(latestPath, "utf-8")) as { formToken: string; writtenAt?: number };
        } catch { return null; }
      })();
      if (!data) {
        return { content: [{ type: "text" as const, text: "no-token" }], structuredContent: { formToken: null, isNew: false } };
      }
      const { formToken, writtenAt } = data;
      // A token is "new" if it differs from what the panel already has AND was written
      // within the last 5 minutes — guards against stale schema files from previous sessions
      // without a timing race against when the panel mounted.
      const FRESH_WINDOW_MS = 5 * 60 * 1000;
      const isFresh = writtenAt === undefined || (Date.now() - writtenAt) < FRESH_WINDOW_MS;
      const isNew = !!formToken && formToken !== panelToken && isFresh;
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
      let schema = readSchema(token) as { teamSiteId: string; libraryContentVersionId: string; templateName?: string } | null;
      if (schema) {
        // Schema is now in the map (readSchema set it). Delete the disk file.
        deleteSchemaFile(token);
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
        writeResult(token, { generatedLivedocId: gid, status: "Generating", downloadUrls: [], templateName: schema?.templateName ?? "" });
        // Reverse-lookup file survives server restarts
        writeGid(gid, token);
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
      dbg(`rawOutputs=${JSON.stringify(outputs)}`);

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
          if (format === "thumbnail") return;  // thumbnail outputs don't have downloadable content
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
          const gidData = readGid(generatedLivedocId);
          if (gidData) formToken = gidData.formToken;
        }
        if (formToken) {
          // Preserve templateName written by submit_form
          let templateName = "";
          const existingResult = readResult(formToken);
          if (existingResult) templateName = existingResult.templateName ?? "";
          writeResult(formToken, { generatedLivedocId, status: "Completed", downloadUrls, downloads, templateName });
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
    "get_preview_images",
    {
      description: "Internal: fetches preview image URLs for a completed generation output. Called by the App panel after generation completes.",
      inputSchema: {
        generatedLivedocId: z.string(),
        outputId: z.string().describe("Output ID or format alias like 'pptx' or 'pdf'"),
      },
      _meta: { ui: { visibility: ["app"] } },
    },
    async (args) => {
      const { generatedLivedocId, outputId } = args as { generatedLivedocId: string; outputId: string };
      const LOG = (msg: string) => { try { fs.appendFileSync(DEBUG_LOG, `[${new Date().toISOString()}] PREV: ${msg}\n`); } catch {} };
      LOG(`gid=${generatedLivedocId} outputId=${outputId}`);
      const res = await apiFetch(`/v3/generatedLivedocs/${generatedLivedocId}/outputs/${outputId}/previewImages`);
      const bodySnip = JSON.stringify(res.body).slice(0, 400);
      LOG(`HTTP ${res.status} body=${bodySnip}`);
      if (res.status !== 200) {
        return {
          content: [{ type: "text" as const, text: `Preview images unavailable: HTTP ${res.status}` }],
          structuredContent: { images: [], httpStatus: res.status },
        };
      }
      const body = res.body as Record<string, unknown>;
      const rawImages = ((body.previewImages ?? body.PreviewImages) as Array<Record<string, unknown>>) ?? [];
      const imagesMeta = rawImages.map(img => ({
        index: Number(img.index ?? img.Index ?? 0),
        url: String(img.url ?? img.Url ?? ""),
      })).filter(img => img.url);
      // Filter to known Seismic domains; the panel iframe loads them directly via
      // the resourceDomains CSP. URLs are now time-expiry signed (no Bearer token needed).
      const images = imagesMeta.filter(img => {
        try {
          const host = new URL(img.url).hostname;
          return host.endsWith(".seismic.com") || host.endsWith(".seismic-dev.com");
        } catch { return false; }
      });
      LOG(`${images.length} images; url[0]=${images[0]?.url?.slice(0,80) ?? "none"}`);
      return {
        content: [{ type: "text" as const, text: `${images.length} preview images` }],
        structuredContent: { images },
      };
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
          const rawUrl = String(
            (body.contentThumbnailImageUrls as string[] | undefined)?.[0] ?? body.imageUrl ?? ""
          );
          if (!rawUrl) return { versionId, thumbnailUrl: "" };
          // Basic allowlist: only pass through URLs on known Seismic domains.
          let parsedHost: string;
          try { parsedHost = new URL(rawUrl).hostname; }
          catch { return { versionId, thumbnailUrl: "" }; }
          const isSeismicHost = parsedHost.endsWith(".seismic.com") || parsedHost.endsWith(".seismic-dev.com");
          if (!isSeismicHost) return { versionId, thumbnailUrl: "" };
          // The App panel iframe can load these URLs directly — the resource registration
          // sets resourceDomains to *.seismic.com / *.seismic-dev.com in the CSP.
          return { versionId, thumbnailUrl: rawUrl };
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
    "login",
    {
      description:
        "Sign in to Seismic. Authentication is handled via the LiveDoc panel UI — do NOT ask the user for credentials in chat. " +
        "If the user needs to sign in, tell them to open the LiveDoc panel where a sign-in form will appear.",
      inputSchema: {
        tenant:   z.string().optional().describe(`Tenant slug, e.g. "qa01eastasia01". Defaults to AUTH_TENANT env var.`),
        username: z.string().optional().describe("Seismic username. Defaults to AUTH_USERNAME env var."),
        password: z.string().optional().describe("Seismic password. Defaults to AUTH_PASSWORD env var."),
      },
      _meta: { ui: { visibility: ["app"] } },
    },
    async (args) => {
      const result = await handleLogin(args as Parameters<typeof handleLogin>[0]);
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    }
  );

  server.registerTool(
    "set_token",
    {
      description: "Internal: manually override the bearer token. Use the LiveDoc panel to sign in instead.",
      inputSchema: {
        token: z.string().describe("The new bearer token (without the 'Bearer ' prefix)."),
      },
      _meta: { ui: { visibility: ["app"] } },
    },
    async (args) => {
      setToken((args as { token: string }).token);
      setTokenManual(true);
      return { content: [{ type: "text" as const, text: "Token updated. This token will not be auto-replaced by the credential-flow login on a 401." }] };
    }
  );
}
