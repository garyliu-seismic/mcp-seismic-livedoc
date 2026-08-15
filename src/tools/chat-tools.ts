import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerAppTool, registerAppResource, RESOURCE_MIME_TYPE } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";
import { handleSearchTemplates, handleSearchContent } from "../handlers/content.js";
import { handleGetInputs, buildFormSchema } from "../handlers/inputs.js";
import { handleSubmitGeneration, handleGetStatus, handleGetDownloadUrl, handleDownloadGenerationOutput } from "../handlers/generation.js";
import { generateToken } from "../utils/os-utils.js";
import { writeSchema, writeLatestToken, readResult, writePrefill, readSchema } from "../ipc/temp-file.js";
import { FORM_RESOURCE_URI } from "../config.js";

export function registerChatTools(server: McpServer): void {
  // MCP App UI resource — served when Claude Desktop opens the App panel.
  // frameDomains CSP is on the registration config (resources/list) so Claude Desktop
  // applies it at connection time. The shell uses callServerTool only — no connectDomains needed.
  registerAppResource(
    server,
    "LiveDoc Form",
    FORM_RESOURCE_URI,
    { description: "LiveDoc input form — React shell built by Vite." } as Parameters<typeof registerAppResource>[3],
    () => {
      const shellPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "dist", "views", "form-shell.html");
      const text = fs.readFileSync(shellPath, "utf-8");
      return {
        contents: [{
          uri: FORM_RESOURCE_URI,
          mimeType: RESOURCE_MIME_TYPE,
          text,
          _meta: {
            ui: {
              csp: {
                // Allow the iframe to load images directly from the Seismic CDN —
                // slide group thumbnails, content candidate thumbnails, and preview images.
                resourceDomains: ["https://*.seismic-dev.com", "https://*.seismic.com"],
              },
            },
          },
        }],
      };
    }
  );

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
        "Do NOT build your own form, do NOT use AskUserQuestion. " +
        "If the user wants sample/default values filled in, call prefill_livedoc_form_values with the formToken — " +
        "do NOT call submit_form, get_form_schema, or submit_livedoc_generation yourself; those bypass the form the user is looking at.",
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

      // External content candidate thumbnails are fetched lazily by the panel via
      // get_candidate_thumbnails (which inlines them server-side). Do NOT pre-fetch
      // them here — there can be 30+ candidates and the per-image timeout would stall
      // get_livedoc_inputs unacceptably.

      writeSchema(formToken, schema);
      // Overwrite the "latest" pointer so the panel can detect a new generation request
      // even when the App panel was already open from a previous run.
      writeLatestToken(formToken);

      // Surface the actual fillable field/table/variable-list names so a subsequent
      // prefill_livedoc_form_values call uses real names instead of guessing generic ones.
      const adhocScalars = (schema.adhocScalars as Array<{ name: string; type: string }>) ?? [];
      const adhocTables = (schema.adhocTables as Array<{ name: string; columns: Array<{ name: string; colType: string }> }>) ?? [];
      const variableLists = (schema.variableLists as Array<{
        name: string;
        scalars: Array<{ name: string; type: string }>;
        tables: Array<{ name: string; columns: Array<{ name: string; colType: string }> }>;
      }>) ?? [];
      const fillableFields = {
        scalars: adhocScalars.map(f => ({ name: f.name, type: f.type })),
        tables: adhocTables.map(t => ({ name: t.name, columns: t.columns.map(c => c.name) })),
        variableLists: variableLists.map(vl => ({
          name: vl.name,
          scalars: vl.scalars.map(f => f.name),
          tables: vl.tables.map(t => ({ name: t.name, columns: t.columns.map(c => c.name) })),
        })),
      };

      return {
        content: [{
          type: "text" as const,
          text: `Form opened in the App panel (formToken="${formToken}"). ` +
            `The user fills it out and clicks Submit in the panel — generation runs automatically. ` +
            `DO NOT call open_form_ui. DO NOT call submit_form. DO NOT call get_form_result. ` +
            `Just tell the user to fill the form. When they say it is done, call get_panel_result with formToken="${formToken}". ` +
            `If sample/default values were requested, call prefill_livedoc_form_values with formToken="${formToken}" — ` +
            `use ONLY the exact field/table/variable-list names listed below, never invent your own: ` +
            JSON.stringify(fillableFields),
        }],
        structuredContent: { formToken, fillableFields },
      };
    }
  );

  server.registerTool(
    "prefill_livedoc_form_values",
    {
      description:
        "Prefill sample or default values into the LiveDoc input form that get_livedoc_inputs already opened in the App panel. " +
        "Use this whenever the user asks for sample/default/suggested values — it writes the values into the SAME open form so the " +
        "user can review and click Submit themselves. Do NOT call submit_form, get_form_schema, or submit_livedoc_generation yourself; " +
        "those submit generation directly and skip the user's review, which is not what 'fill in sample values' means. " +
        "Only include field/table/variable-list names that exist in the schema fields you saw from get_livedoc_inputs.",
      inputSchema: {
        formToken: z.string().describe("The formToken returned by get_livedoc_inputs."),
        scalars: z.record(z.string(), z.any()).optional().describe("Map of adhoc scalar field name -> suggested value."),
        tables: z.record(z.string(), z.array(z.record(z.string(), z.any()))).optional().describe("Map of adhoc table name -> array of row objects keyed by column name."),
        variableLists: z.record(z.string(), z.object({
          scalars: z.record(z.string(), z.any()).optional().describe("Map of scalar field name -> suggested value."),
          tables: z.record(z.string(), z.array(z.record(z.string(), z.any()))).optional().describe("Map of table name -> array of row objects keyed by column name."),
        })).optional().describe("Map of variable list name -> its scalar/table values."),
      },
    },
    async (args) => {
      const { formToken, scalars, tables, variableLists } = args as {
        formToken: string;
        scalars?: Record<string, unknown>;
        tables?: Record<string, Array<Record<string, unknown>>>;
        variableLists?: Record<string, { scalars?: Record<string, unknown>; tables?: Record<string, Array<Record<string, unknown>>> }>;
      };

      const schema = readSchema(formToken) as {
        adhocScalars?: Array<{ name: string }>;
        adhocTables?: Array<{ name: string }>;
        variableLists?: Array<{ name: string }>;
      } | null;
      if (!schema) {
        return {
          content: [{ type: "text" as const, text: "error: no form open for this formToken (schema not found)" }],
          structuredContent: { error: "Schema not found for token" },
          isError: true,
        };
      }

      const validScalars = new Set((schema.adhocScalars ?? []).map(f => f.name));
      const validTables = new Set((schema.adhocTables ?? []).map(t => t.name));
      const validVLs = new Set((schema.variableLists ?? []).map(vl => vl.name));
      const unknown: string[] = [];
      Object.keys(scalars ?? {}).forEach(k => { if (!validScalars.has(k)) unknown.push(`scalars.${k}`); });
      Object.keys(tables ?? {}).forEach(k => { if (!validTables.has(k)) unknown.push(`tables.${k}`); });
      Object.keys(variableLists ?? {}).forEach(k => { if (!validVLs.has(k)) unknown.push(`variableLists.${k}`); });
      if (unknown.length > 0) {
        return {
          content: [{
            type: "text" as const,
            text: `error: unknown field name(s): ${unknown.join(", ")}. ` +
              `Valid names — scalars: ${[...validScalars].join(", ") || "(none)"}; ` +
              `tables: ${[...validTables].join(", ") || "(none)"}; ` +
              `variableLists: ${[...validVLs].join(", ") || "(none)"}. ` +
              `Re-call prefill_livedoc_form_values using only these exact names.`,
          }],
          structuredContent: { error: "Unknown field name(s)", unknown },
          isError: true,
        };
      }

      writePrefill(formToken, { scalars, tables, variableLists });
      return {
        content: [{
          type: "text" as const,
          text: "Sample values sent to the open form. Tell the user their fields have been filled in and ask them to review and click Submit in the panel.",
        }],
        structuredContent: { ok: true },
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
        "and a small 'Links expire â€¦' note at the bottom. Keep the HTML concise (no external resources). " +
        "(2) Also call download_generation_output for each output to save the files locally. " +
        "If status is still 'Generating', tell the user to wait and offer to check again.",
      inputSchema: {
        formToken: z.string().describe("The formToken returned by get_livedoc_inputs."),
      },
    },
    async (args) => {
      const { formToken } = args as { formToken: string };
      const data = readResult(formToken);
      if (!data) {
        return {
          content: [{ type: "text" as const, text: "No result yet — generation has not started or the form has not been submitted. Check the App panel." }],
        };
      }
      const text = data.status === "Completed"
        ? `Generation complete. templateName: "${data.templateName ?? ""}". generatedLivedocId: ${data.generatedLivedocId}. ` +
          `downloads: ${JSON.stringify(data.downloads ?? [])}. ` +
          `Create an HTML artifact showing the result card, then call download_generation_output for each output.`
        : `Generation status: ${data.status}. generatedLivedocId: ${data.generatedLivedocId}. Check back shortly.`;
      return {
        content: [{ type: "text" as const, text }],
        structuredContent: data as unknown as Record<string, unknown>,
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
          text: `âœ… **${dlFile}** is ready.\n\nDownload link: [${dlFile}](${dlUrl})\n\n(Reproduce the markdown link above verbatim in your reply so the user can click it.)`,
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
}
