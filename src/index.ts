#!/usr/bin/env node
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
} from "@modelcontextprotocol/sdk/types.js";
import * as http from "http";
import { randomUUID } from "crypto";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { exec } from "child_process";

const BASE_URL = process.env.SEISMIC_BASE_URL ?? "https://api.seismic.com/livedoc";

const DEFAULT_AUTH_URI      = process.env.AUTH_SERVICE_URI    ?? "";
const DEFAULT_AUTH_TENANT   = process.env.AUTH_TENANT         ?? "";
const DEFAULT_CLIENT_ID     = process.env.AUTH_CLIENT_ID      ?? "";
const DEFAULT_CLIENT_SECRET = process.env.AUTH_CLIENT_SECRET  ?? "";
const DEFAULT_USERNAME      = process.env.AUTH_USERNAME        ?? "";
const DEFAULT_PASSWORD      = process.env.AUTH_PASSWORD        ?? "";

// ── Tiny HTTP server for artifact form submissions ──────────────────────────
const FORM_PORT = 3099;
const pendingForms = new Map<string, (payload: unknown) => void>();

const formHttpServer = http.createServer((req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }
  const m = req.url?.match(/^\/submit\/([^/?]+)/);
  if (req.method === "POST" && m) {
    const token = m[1];
    let body = "";
    req.on("data", (chunk: Buffer) => { body += chunk; });
    req.on("end", () => {
      const resolver = pendingForms.get(token);
      if (resolver) {
        pendingForms.delete(token);
        try { resolver(JSON.parse(body)); } catch { resolver(body); }
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ ok: true }));
      } else {
        res.writeHead(410); res.end("expired");
      }
    });
  } else {
    res.writeHead(404); res.end();
  }
});
formHttpServer.listen(FORM_PORT, "127.0.0.1");
formHttpServer.on("error", () => { /* port in use — form falls back to copy-paste */ });

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

const tools: Tool[] = [
  {
    name: "search_livedoc_templates",
    description:
      "Search for LiveDoc (Document Generator) templates in Seismic by name or keyword. Returns contentVersionId and teamSiteId needed for other tools. " +
      "ALWAYS call this FIRST whenever the user names or describes a template (e.g. \"generate ContentSelectorForm\") — do NOT ask the user for teamSiteId/libraryContentVersionId directly; those are internal ids the user is unlikely to know. Only ask the user to disambiguate if this search returns zero or multiple plausible matches.",
    inputSchema: {
      type: "object",
      properties: {
        searchText: {
          type: "string",
          description: "Text to search across template title, description, and body.",
        },
        page_size: {
          type: "number",
          description: "Number of results to return (default 10, max 50).",
          default: 10,
        },
      },
      required: [],
    },
  },
  {
    name: "search_livedoc_content",
    description:
      "Search the Seismic library for content items by keyword. Use this when a LiveDoc template has external content inputs (e.g. contentType 'ExternalSlides', 'LiveSlide', 'ExternalStaticSlides') — search for the real document the user wants to include, then use the returned contentVersionId as the id in the submission payload.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search text (title or keyword).",
        },
        contentType: {
          type: "string",
          enum: ["ExternalSlides", "LiveSlide", "ExternalStaticSlides", "LiveDoc", "PDF"],
          description: "Filter by content type. Pass the contentType value exactly as it appears in the template's manualSelectContentItems.",
        },
        page_size: {
          type: "number",
          description: "Number of results to return (default 10, max 50).",
          default: 10,
        },
      },
      required: ["query"],
    },
  },
  {
    name: "get_livedoc_inputs",
    description:
      "Retrieve the full input schema for a LiveDoc template version — ad hoc variables, variable lists, image placeholders, manual select groups, and available output forms. Call this before submitting generation to know what inputs are required. IMPORTANT: After this tool returns, you MUST immediately present a friendly input form to the user: show each adHocInputDefinition as a labelled field (use AskUserQuestion for fields with predefined choices, a markdown table for free-text fields). Do not ask the user whether to show the form — just show it.",
    inputSchema: {
      type: "object",
      properties: {
        teamSiteId: {
          type: "string",
          description: "Team site identifier (UUID) that owns the template.",
        },
        libraryContentVersionId: {
          type: "string",
          description: "Content version identifier (UUID) of the LiveDoc template.",
        },
      },
      required: ["teamSiteId", "libraryContentVersionId"],
    },
  },
  {
    name: "submit_livedoc_generation",
    description:
      "Submit a LiveDoc generation job. Provide ad hoc input values and at least one output format (PPTX, DOCX, PDF). Returns a generatedLivedocId to poll for status.",
    inputSchema: {
      type: "object",
      properties: {
        teamSiteId: {
          type: "string",
          description: "Team site identifier (UUID).",
        },
        libraryContentVersionId: {
          type: "string",
          description: "Content version identifier (UUID) of the LiveDoc template.",
        },
        adHocInputs: {
          type: "array",
          description: "Array of {name, value} pairs for ALL ad hoc inputs — both scalar and table. Scalar value is a primitive (string/number/boolean/ISO-date). Table value is {columns: [\"col1\", ...], rows: [[row1val1, ...], [row2val1, ...]]}.",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              value: {},
            },
            required: ["name", "value"],
          },
        },
        outputs: {
          type: "array",
          description:
            'Output formats to generate. Each item needs "format" (e.g. "PPTX", "DOCX", "PDF"). ALWAYS also set "fileName" — derive it from the template name (e.g. templateName + "." + format.toLowerCase()) — because the API leaves it blank in get_generation_status/download when omitted, producing an unhelpful generic filename after download.',
          items: {
            type: "object",
            properties: {
              format: { type: "string" },
              name: { type: "string" },
              fileName: { type: "string", description: "Filename with extension for the downloaded file, e.g. \"Form Smoke Test_4.pdf\". Always set this." },
            },
            required: ["format", "fileName"],
          },
        },
        variableListData: {
          type: "array",
          description: "Variable list data from variableListDefinitions. Each entry has variableListName and variableInputs. Each variable input value can be a scalar (string/number/boolean/ISO-date) OR a table using {columns: [...], rows: [[...], ...]}.",
          items: {
            type: "object",
            properties: {
              variableListName: { type: "string" },
              variableInputs: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    name: { type: "string" },
                    value: {},
                  },
                  required: ["name", "value"],
                },
              },
            },
            required: ["variableListName", "variableInputs"],
          },
        },
        liveFormSellerTemplateId: {
          type: "string",
          description: "Optional seller template ID to constrain inputs.",
        },
        regionalFormat: {
          type: "string",
          description: 'Optional regional format culture name, e.g. "en-US".',
        },
        manualSelectContentInput: {
          type: "object",
          description:
            'Content selection (from get_livedoc_inputs\' ManualSelectContentInput.ManualSelectContentItems). Everything is included by default - only pass items you want to EXCLUDE, with isInclude: false, or items that need a resolved versionId (see below). ' +
            'IMPORTANT: "id" is a stable SLOT identifier from get_livedoc_inputs — always echo it back UNCHANGED, never replace it with a resolved content id. ' +
            'For contentType "LiveSlide", "ExternalStaticSlides", "ResourcePDF", or "ResourcePDFPage": these slots need a real piece of content assigned. Call search_livedoc_content(query: item.name, contentType: item.contentType) to find it, then set "versionId" to the result\'s contentVersionId (and "sourceBlobId" too if the search result provides a distinct source blob id — needed for ExternalStaticSlides/slide-source references). For "ResourcePDFPage", also set "pageNumber". Do NOT invent a versionId — if search_livedoc_content finds nothing, leave the item out or set isInclude:false rather than fabricating one.',
          properties: {
            manualSelectContentItems: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  id: { type: "string", description: "Stable slot identifier — copy verbatim from get_livedoc_inputs, never replace." },
                  name: { type: "string" },
                  contentType: {
                    type: "string",
                    description:
                      'One of "Group", "Section", "LiveSlide", "ExternalStaticSlides", "ResourcePDF", "ResourcePDFPage" - copy from the matching item in get_livedoc_inputs.',
                  },
                  versionId: {
                    type: "string",
                    description:
                      "Resolved contentVersionId from search_livedoc_content. Required for LiveSlide/ExternalStaticSlides/ResourcePDF/ResourcePDFPage items being included — never fabricate this value.",
                  },
                  sourceBlobId: {
                    type: "string",
                    description: "Resolved source blob id, when the search result provides one distinct from versionId (e.g. ExternalStaticSlides).",
                  },
                  pageNumber: {
                    type: "number",
                    description: "Page number within the resource, for contentType ResourcePDFPage only.",
                  },
                  isInclude: { type: "boolean", description: "Set false to exclude this item. Defaults to true." },
                  orderIndex: { type: "number" },
                },
                required: ["id", "contentType", "isInclude"],
              },
            },
          },
        },
      },
      required: ["teamSiteId", "libraryContentVersionId", "adHocInputs", "outputs"],
    },
  },
  {
    name: "get_generation_status",
    description:
      "Check the status of a LiveDoc generation job. Poll until all outputs reach 'Completed' or 'Failed'.",
    inputSchema: {
      type: "object",
      properties: {
        generatedLivedocId: {
          type: "string",
          description: "The generatedLivedocId returned by submit_livedoc_generation.",
        },
      },
      required: ["generatedLivedocId"],
    },
  },
  {
    name: "open_form_ui",
    description:
      "Returns a deep-link URL that opens the LiveDoc Form UI for a complex template. Call this when get_livedoc_inputs returns isComplex=true. Optionally pass prefillValues to pre-populate form fields with AI-suggested defaults — the user can review and edit before generating. The form UI renders all widgets (tables, image uploads, slide selectors) and handles generation and download.",
    inputSchema: {
      type: "object",
      properties: {
        teamSiteId: { type: "string", description: "Team site identifier (UUID)." },
        libraryContentVersionId: { type: "string", description: "Content version identifier (UUID) of the LiveDoc template." },
        context: { type: "string", description: "The user's original generation request (natural language). Shown as a hint in the form." },
        prefillValues: {
          type: "object",
          description: "Optional AI-suggested default values to pre-populate the form. Shape matches the generate request body. User can review and override before submitting.",
          properties: {
            adHocInputs: {
              type: "array",
              items: { type: "object", properties: { name: { type: "string" }, value: {} }, required: ["name", "value"] },
            },
            variableListData: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  variableListName: { type: "string" },
                  variableInputs: {
                    type: "array",
                    items: { type: "object", properties: { name: { type: "string" }, value: {} }, required: ["name", "value"] },
                  },
                },
                required: ["variableListName", "variableInputs"],
              },
            },
          },
        },
      },
      required: ["teamSiteId", "libraryContentVersionId"],
    },
  },
  {
    name: "wait_for_form_submit",
    description:
      "Wait (up to 10 minutes) for the user to fill and submit the LiveDoc input form rendered in the HTML artifact. Returns the form payload — pass it directly to submit_livedoc_generation. Call this IMMEDIATELY after creating the artifact, without waiting for the user first.",
    inputSchema: {
      type: "object",
      properties: {
        token: {
          type: "string",
          description: "The form session token returned by get_livedoc_inputs.",
        },
      },
      required: ["token"],
    },
  },
  {
    name: "get_form_result",
    description:
      "Retrieve the generation result posted back by the Form UI after the user completed and closed the form. Call this after the user confirms the form tab closed. Returns generatedLivedocId and all output statuses — then call get_generation_download_url for each completed output to get download links.",
    inputSchema: {
      type: "object",
      properties: {
        token: {
          type: "string",
          description: "The token returned by open_form_ui.",
        },
      },
      required: ["token"],
    },
  },
  {
    name: "login",
    description:
      "Obtain a Seismic bearer token using username and password (OAuth 2.0 Resource Owner Password Credentials). Automatically sets the token for all subsequent tool calls. Use this instead of set_token when you have credentials rather than a pre-issued token.",
    inputSchema: {
      type: "object",
      properties: {
        username: { type: "string", description: `Seismic username. Defaults to AUTH_USERNAME env var${DEFAULT_USERNAME ? " (pre-configured)" : ""}.` },
        password: { type: "string", description: `Seismic password. Defaults to AUTH_PASSWORD env var${DEFAULT_PASSWORD ? " (pre-configured)" : ""}.` },
        tenant:   { type: "string", description: `Tenant slug, e.g. "qa01eastasia01". Defaults to AUTH_TENANT env var (${DEFAULT_AUTH_TENANT || "not set"}).` },
        authServiceUri: { type: "string", description: `Auth service base URL. Defaults to AUTH_SERVICE_URI env var (${DEFAULT_AUTH_URI || "not set"}).` },
        clientId:     { type: "string", description: "OAuth client ID. Defaults to AUTH_CLIENT_ID env var." },
        clientSecret: { type: "string", description: "OAuth client secret. Defaults to AUTH_CLIENT_SECRET env var." },
      },
      required: [],
    },
  },
  {
    name: "set_token",
    description:
      "Update the Seismic API bearer token used by all other tools. Call this whenever a tool returns HTTP 401 before retrying.",
    inputSchema: {
      type: "object",
      properties: {
        token: {
          type: "string",
          description: "The new bearer token (without the 'Bearer ' prefix).",
        },
      },
      required: ["token"],
    },
  },
  {
    name: "get_generation_download_url",
    description:
      "Get the download URL for a completed LiveDoc output. Returns a JSON payload with the URL (does not redirect). Use outputId from get_generation_status, or use a format alias like 'pptx', 'docx', 'pdf'.",
    inputSchema: {
      type: "object",
      properties: {
        generatedLivedocId: {
          type: "string",
          description: "The generatedLivedocId.",
        },
        outputId: {
          type: "string",
          description:
            "The output id from get_generation_status, or a format alias: 'pptx', 'docx', 'pdf', 'gslides', 'gdoc'.",
        },
      },
      required: ["generatedLivedocId", "outputId"],
    },
  },
  {
    name: "download_generation_output",
    description:
      "Download a completed LiveDoc output to a local file (the OS Downloads folder) and open it with the system default app " +
      "(e.g. PowerPoint for PPTX, Word for DOCX) — the same effect as clicking a browser download and double-clicking the file. " +
      "Use this when the user wants to view/inspect the generated document directly instead of just getting a link. " +
      "Use outputId from get_generation_status, or a format alias like 'pptx', 'docx', 'pdf'.",
    inputSchema: {
      type: "object",
      properties: {
        generatedLivedocId: {
          type: "string",
          description: "The generatedLivedocId.",
        },
        outputId: {
          type: "string",
          description:
            "The output id from get_generation_status, or a format alias: 'pptx', 'docx', 'pdf', 'gslides', 'gdoc'.",
        },
        autoOpen: {
          type: "boolean",
          description: "Whether to automatically open the file with the OS default app after downloading. Default true.",
          default: true,
        },
      },
      required: ["generatedLivedocId", "outputId"],
    },
  },

  // ── PPTX Auto-Tagging tools (PoC) ─────────────────────────────────────────
  {
    name: "pptx_extract_shapes",
    description:
      "Extract all shapes from a PPTX file and return a structured list with shape ID, type (text/table/chart/image), name, alt-text, and text content. " +
      "Use this first to understand what's in the template before deciding which shapes to mark as dynamic. " +
      "Requires the local PoC server to be running (cd poc-auto-tagging && npm start).",
    inputSchema: {
      type: "object",
      properties: {
        pptxBase64: {
          type: "string",
          description: "Base64-encoded PPTX file content.",
        },
        slideIndex: {
          type: "number",
          description: "Optional: only extract shapes from this slide (0-based). Omit to extract all slides.",
        },
      },
      required: ["pptxBase64"],
    },
  },
  {
    name: "pptx_auto_tag",
    description:
      "Send a PPTX to the AI (local LLM) for automatic analysis. The AI identifies which shapes should be dynamic " +
      "and suggests variable names and types. Returns suggestions with confidence scores. " +
      "Optionally provide a datasource schema to improve matching accuracy. " +
      "Requires the local PoC server to be running (cd poc-auto-tagging && npm start).",
    inputSchema: {
      type: "object",
      properties: {
        pptxBase64: {
          type: "string",
          description: "Base64-encoded PPTX file content.",
        },
        schema: {
          type: "object",
          description: "Optional datasource schema as a JSON object, e.g. { \"companyName\": \"string\", \"revenue\": \"number\" }. Providing this greatly improves accuracy.",
          additionalProperties: { type: "string" },
        },
      },
      required: ["pptxBase64"],
    },
  },
  {
    name: "pptx_mark_shapes",
    description:
      "Apply dynamic element markings to a PPTX file. For each mark: writes [[varName]] to the shape's alt-text " +
      "and injects a stable GUID into p:tags. Returns the marked PPTX as base64 and a binding manifest. " +
      "Requires the local PoC server to be running (cd poc-auto-tagging && npm start).",
    inputSchema: {
      type: "object",
      properties: {
        pptxBase64: {
          type: "string",
          description: "Base64-encoded PPTX file content.",
        },
        marks: {
          type: "array",
          description: "List of shapes to mark as dynamic.",
          items: {
            type: "object",
            properties: {
              slideIndex: { type: "number", description: "0-based slide index." },
              shapeId:    { type: "number", description: "Numeric shape ID (cNvPr/@id) from pptx_extract_shapes." },
              varName:    { type: "string", description: "camelCase variable name, e.g. 'companyName'." },
              varType:    { type: "string", enum: ["text", "image", "table", "chart", "number", "date"], description: "Type of dynamic element." },
              description: { type: "string", description: "Optional: human-readable description of what this variable represents." },
            },
            required: ["slideIndex", "shapeId", "varName", "varType"],
          },
        },
      },
      required: ["pptxBase64", "marks"],
    },
  },
  {
    name: "pptx_get_manifest",
    description:
      "Extract the current binding manifest from a PPTX file — lists all shapes already marked as dynamic " +
      "(via alt-text [[varName]] convention). Useful for verifying what has been marked. " +
      "Requires the local PoC server to be running (cd poc-auto-tagging && npm start).",
    inputSchema: {
      type: "object",
      properties: {
        pptxBase64: {
          type: "string",
          description: "Base64-encoded PPTX file content.",
        },
      },
      required: ["pptxBase64"],
    },
  },
];

// ── Form HTML builder ───────────────────────────────────────────────────────

function esc(s: string): string {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function gf(o: Record<string, unknown>, key: string): unknown {
  return o[key] ?? o[key[0].toUpperCase() + key.slice(1)];
}

function typeLabel(type: string): string {
  switch (type.toUpperCase()) {
    case "STRING":  return " (string)";
    case "INTEGER": return " (integer)";
    case "FLOAT":   return " (float)";
    case "DATE":    return " (date)";
    case "BOOL": case "BOOLEAN": return " (boolean)";
    default: return "";
  }
}

function scalarInput(name: string, type: string, scope: string, vlName?: string): string {
  const t = type.toUpperCase();
  const id = "f-" + name.replace(/[^a-zA-Z0-9]/g, "_");
  const vlAttr = vlName ? ` data-vl-name="${esc(vlName)}"` : "";
  const label = esc(name) + typeLabel(t);
  if (t === "BOOL" || t === "BOOLEAN") {
    return `<div class="bool-field"><input type="checkbox" id="${id}" data-scope="${scope}" data-field-name="${esc(name)}" data-field-type="${esc(t)}"${vlAttr}><label for="${id}">${label}</label></div>`;
  }
  const itype = t === "DATE" ? "date" : (t === "INTEGER" || t === "FLOAT") ? "number" : "text";
  const step = t === "FLOAT" ? ` step="any"` : t === "INTEGER" ? ` step="1"` : "";
  return `<div class="fw"><label class="fl" for="${id}">${label}</label><input type="${itype}"${step} id="${id}" class="fi" data-scope="${scope}" data-field-name="${esc(name)}" data-field-type="${esc(t)}"${vlAttr} placeholder="${esc(name)}"></div>`;
}

function tableInput(name: string, columns: Array<Record<string, unknown>>, scope: string, vlName?: string): string {
  const tid = "tbl-" + name.replace(/[^a-zA-Z0-9]/g, "_");
  const colDefs = columns.map(c => ({ name: String(gf(c, "name") ?? ""), type: String(gf(c, "type") ?? "STRING") }));
  const colsAttr = esc(JSON.stringify(colDefs));
  const vlAttr = vlName ? ` data-vl-name="${esc(vlName)}"` : "";
  const ths = colDefs.map(c => `<th>${esc(c.name)}</th>`).join("") + `<th style="width:32px"></th>`;
  return `<div class="tbl-wrap"><div class="sl">${esc(name)}</div><table class="dt" data-table-id="${tid}" data-table-scope="${scope}" data-table-name="${esc(name)}"${vlAttr} data-cols="${colsAttr}"><thead><tr>${ths}</tr></thead><tbody id="${tid}-body"></tbody></table><button class="add-btn" onclick="addRow('${tid}')">+ Add row</button></div>`;
}

function buildFormHtml(
  templateName: string,
  adhocInputs: Array<Record<string, unknown>>,
  varListInputs: Array<Record<string, unknown>>,
  manualSelect: Record<string, unknown> | undefined,
  forms: Array<Record<string, unknown>>,
  teamSiteId: string,
  versionId: string,
  token: string
): string {
  const isTable = (i: Record<string, unknown>) => ((gf(i, "columns") as unknown[] | undefined)?.length ?? 0) > 0;
  const scalarAdhoc = adhocInputs.filter(i => !isTable(i));
  const tableAdhoc  = adhocInputs.filter(i => isTable(i));

  const scalarGrid = scalarAdhoc.length
    ? `<div class="grid">${scalarAdhoc.map(i => scalarInput(String(gf(i, "name") ?? ""), String(gf(i, "type") ?? "STRING"), "adhoc")).join("")}</div>`
    : "";

  const tableHtml = tableAdhoc.map(i =>
    tableInput(String(gf(i, "name") ?? ""), (gf(i, "columns") as Array<Record<string, unknown>>) ?? [], "adhoc")
  ).join("");

  const vlHtml = varListInputs.map(vl => {
    const vlName   = String(gf(vl, "variableListName") ?? "");
    const dsName   = String(gf(vl, "dataSourceName") ?? gf(vl, "dataSourceId") ?? "");
    const inputs   = (gf(vl, "variableInputs") as Array<Record<string, unknown>>) ?? [];
    const scVl     = inputs.filter(i => !isTable(i));
    const tVl      = inputs.filter(i => isTable(i));
    const dsSpan   = dsName ? ` <span class="badge">${esc(dsName)}</span>` : "";
    const scGrid   = scVl.length ? `<div class="grid">${scVl.map(i => scalarInput(String(gf(i, "name") ?? ""), String(gf(i, "type") ?? "STRING"), "vl", vlName)).join("")}</div>` : "";
    const tblParts = tVl.map(i => tableInput(String(gf(i, "name") ?? ""), (gf(i, "columns") as Array<Record<string, unknown>>) ?? [], "vl", vlName)).join("");
    return `<div class="section"><div class="sl">Variable list — ${esc(vlName)}${dsSpan}</div>${scGrid}${tblParts}</div>`;
  }).join("");

  // Manual select groups/sections (plain include/exclude) and external content (one or more
  // resolved candidates picked via checkboxes — candidates are pre-attached by handleGetInputs).
  let msHtml = "";
  if (manualSelect) {
    const items = (gf(manualSelect, "manualSelectContentItems") as Array<Record<string, unknown>>) ?? [];
    const groups = items.filter(i => ["Group", "Section"].includes(String(gf(i, "contentType") ?? "")));
    const external = items.filter(i => !["Group", "Section"].includes(String(gf(i, "contentType") ?? "")));

    const checks = groups.map(g => {
      const gId = esc(String(gf(g, "id") ?? ""));
      const gName = esc(String(gf(g, "name") ?? ""));
      const gType = esc(String(gf(g, "contentType") ?? "Group"));
      const inc = gf(g, "isInclude") !== false ? " checked" : "";
      const oi = Number(gf(g, "orderIndex") ?? 0);
      return `<label class="grp"><input type="checkbox"${inc} data-group-id="${gId}" data-group-name="${gName}" data-order-index="${oi}" data-content-type="${gType}"> <span>${gName}</span></label>`;
    }).join("");
    const groupsHtml = groups.length
      ? `<div class="section"><div class="sl">Content selection</div><div class="grp-list">${checks}</div></div>`
      : "";

    const externalHtml = external.map(item => {
      const iId = esc(String(gf(item, "id") ?? ""));
      const iName = esc(String(gf(item, "name") ?? ""));
      const oi = Number(gf(item, "orderIndex") ?? 0);
      const candidates = (gf(item, "candidates") as Array<Record<string, unknown>> | undefined) ?? [];
      if (!candidates.length) {
        return `<div class="ext-item"><div class="sl" style="margin-bottom:6px">${iName}</div><span class="badge" style="background:#fde8e8;color:#c00">No matching content found</span></div>`;
      }
      const totalCount = Number(gf(item, "candidatesTotalCount") ?? candidates.length);
      const truncatedNote = totalCount > candidates.length
        ? ` <span class="badge">Showing ${candidates.length} of ${totalCount} matches — refine the template's content filter if you need a different one</span>`
        : "";
      // Multiple documents can be attached to the same slot, so each candidate is its own
      // checkbox rather than a single-select dropdown — checking N boxes submits N items
      // that all share this slot's id/name but carry different resolved versionId/format.
      const candidateChecks = candidates.map((c, i) => {
        const val = esc(JSON.stringify({ versionId: gf(c, "versionId"), sourceBlobId: gf(c, "sourceBlobId"), format: gf(c, "format") }));
        const label = esc(`${String(gf(c, "title") ?? "")} (${String(gf(c, "format") ?? "")})`);
        const checkedAttr = i === 0 ? " checked" : "";
        return `<label class="grp"><input type="checkbox"${checkedAttr} data-external-candidate="${iId}" data-external-name="${iName}" data-order-index="${oi}" value='${val}'> <span>${label}</span></label>`;
      }).join("");
      return `<div class="ext-item"><div class="sl" style="margin-bottom:6px">${iName}${truncatedNote}</div><div class="grp-list">${candidateChecks}</div></div>`;
    }).join("");

    msHtml = groupsHtml + (externalHtml ? `<div class="section"><div class="sl">External content</div>${externalHtml}</div>` : "");
  }

  // Group forms by unique name → { formName: [{outputs}, ...] }
  type FormCfg = { outputs: Array<{ format: unknown }> };
  const formsByName = new Map<string, FormCfg[]>();
  for (const f of forms) {
    const name = String(gf(f, "name") ?? "");
    if (!formsByName.has(name)) formsByName.set(name, []);
    formsByName.get(name)!.push({
      outputs: ((gf(f, "outputs") as Array<Record<string, unknown>>) ?? []).map(o => ({ format: gf(o, "format") })),
    });
  }
  const uniqueFormNames = Array.from(formsByName.keys());
  const multiForm = uniqueFormNames.length > 1;

  // Initial state: first form name, first output combo
  const firstFormName = uniqueFormNames[0] ?? "";
  const firstFormCfgs = formsByName.get(firstFormName) ?? [];
  const initOutputs = JSON.stringify(firstFormCfgs[0]?.outputs ?? []);

  // Form selector (only when >1 distinct form name)
  const formSelHtml = multiForm
    ? `<div class="fmt-row"><div class="sl" style="margin-bottom:10px">Select form</div><div id="form-sel">${
        uniqueFormNames.map((n, i) =>
          `<button class="fmt${i === 0 ? " active" : ""}" data-form-name="${esc(n)}" onclick="selForm(this)">${esc(n)}</button>`
        ).join("")
      }</div></div>`
    : "";

  // Initial output format buttons (format codes, not form names)
  const fmtBtnsHtml = firstFormCfgs.map((cfg, i) => {
    const label = esc(cfg.outputs.map(o => String(o.format)).join(" + ") || "Default");
    return `<button class="fmt${i === 0 ? " active" : ""}" data-outputs="${esc(JSON.stringify(cfg.outputs))}" onclick="selFmt(this)">${label}</button>`;
  }).join("");

  // JS config: all form configs keyed by name
  const formConfigsJs = JSON.stringify(
    Object.fromEntries(Array.from(formsByName.entries()))
  );

  const css = `*{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
body{background:#fff;padding:20px;font-size:14px;color:#1d1d1f}
.title{font-size:17px;font-weight:700;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px 16px;margin-bottom:20px;align-items:end}
.fw{display:flex;flex-direction:column;gap:5px}
.fl{font-size:12px;font-weight:600;color:#555}
.fi{width:100%;padding:7px 10px;border:1px solid #d0d0d0;border-radius:7px;font-size:13px;outline:none;background:#fff}
.fi:focus{border-color:#0066cc}
.bool-field{display:flex;align-items:center;gap:8px;padding-bottom:10px}
.bool-field input{width:16px;height:16px;cursor:pointer;flex-shrink:0}
.bool-field label{font-size:12px;font-weight:600;color:#555;cursor:pointer}
.section{margin-bottom:20px}
.sl{font-size:13px;font-weight:700;color:#444;margin-bottom:10px}
.badge{display:inline-block;padding:1px 7px;border-radius:100px;font-size:11px;font-weight:600;background:#e8f0fe;color:#0066cc;margin-left:6px}
.tbl-wrap{margin-bottom:20px}
.dt{width:100%;border-collapse:collapse;font-size:13px}
.dt th{font-size:12px;font-weight:600;color:#666;padding:6px 8px;text-align:left;border-bottom:1px solid #e5e5e5;background:#fafafa}
.dt td{padding:4px 6px;border-bottom:1px solid #f5f5f5}
.dt td input[type=text],.dt td input[type=number],.dt td input[type=date]{width:100%;padding:5px 7px;border:1px solid #d0d0d0;border-radius:5px;font-size:12px;outline:none}
.dt td input[type=text]:focus,.dt td input[type=number]:focus,.dt td input[type=date]:focus{border-color:#0066cc}
.dt td input[type=checkbox]{width:16px;height:16px;cursor:pointer}
.del{background:none;border:none;cursor:pointer;color:#ccc;font-size:15px;padding:2px 5px}
.del:hover{color:#c00}
.add-btn{font-size:12px;color:#0066cc;background:none;border:1px dashed #0066cc;border-radius:6px;padding:5px 14px;cursor:pointer;margin-top:6px}
.add-btn:hover{background:#e8f0fe}
.fmt-row{margin-bottom:20px}
.fmt{padding:7px 18px;border-radius:20px;border:none;background:#f0f0f0;color:#444;font-size:13px;font-weight:500;cursor:pointer;margin-right:8px;transition:background .15s}
.fmt.active{background:#0066cc;color:#fff}
.sub{display:inline-flex;align-items:center;gap:6px;padding:9px 22px;background:#0066cc;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer}
.sub:hover{background:#0055b3}
.grp-list{display:flex;flex-direction:column;gap:8px}
.grp{display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px}
.grp input{width:16px;height:16px}
.ext-item{margin-bottom:14px;padding:10px 12px;border:1px solid #e5e5e5;border-radius:8px}`;

  const js = `var sel=${initOutputs};
var tsid=${JSON.stringify(teamSiteId)};
var vid=${JSON.stringify(versionId)};
var formToken=${JSON.stringify(token)};
var formCfgs=${formConfigsJs};
function selFmt(b){
  sel=JSON.parse(b.dataset.outputs);
  document.querySelectorAll('#fmt-btns .fmt').forEach(function(x){x.classList.remove('active')});
  b.classList.add('active');
}
function selForm(b){
  document.querySelectorAll('#form-sel .fmt').forEach(function(x){x.classList.remove('active')});
  b.classList.add('active');
  var cfgs=formCfgs[b.dataset.formName]||[];
  var div=document.getElementById('fmt-btns');
  div.innerHTML='';
  cfgs.forEach(function(cfg,i){
    var btn=document.createElement('button');
    btn.className='fmt'+(i===0?' active':'');
    btn.dataset.outputs=JSON.stringify(cfg.outputs);
    btn.textContent=cfg.outputs.map(function(o){return o.format;}).join(' + ')||'Default';
    btn.onclick=function(){selFmt(this);};
    div.appendChild(btn);
  });
  if(cfgs.length)sel=cfgs[0].outputs;
}
function addRow(tid){
  var tbl=document.querySelector('[data-table-id="'+tid+'"]');
  if(!tbl)return;
  var cols=JSON.parse(tbl.dataset.cols||'[]');
  var tb=document.getElementById(tid+'-body');
  if(!tb)return;
  var tr=document.createElement('tr');
  cols.forEach(function(c){
    var td=document.createElement('td');
    var t=(c.type||'').toUpperCase();
    var el;
    if(t==='BOOL'||t==='BOOLEAN'){el=document.createElement('input');el.type='checkbox';}
    else if(t==='DATE'){el=document.createElement('input');el.type='date';}
    else{el=document.createElement('input');el.type=(t==='INTEGER'||t==='FLOAT')?'number':'text';if(t==='INTEGER')el.step='1';if(t==='FLOAT')el.step='any';}
    el.dataset.colType=t;td.appendChild(el);tr.appendChild(td);
  });
  var dtd=document.createElement('td');
  var db=document.createElement('button');
  db.className='del';db.innerHTML='&#128465;';
  db.onclick=function(){this.closest('tr').remove();};
  dtd.appendChild(db);tr.appendChild(dtd);tb.appendChild(tr);
}
function tblVal(tbl){
  var cols=JSON.parse(tbl.dataset.cols||'[]');
  var rows=Array.from(tbl.querySelectorAll('tbody tr')).map(function(tr){
    return Array.from(tr.querySelectorAll('[data-col-type]')).map(function(c){
      var t=(c.dataset.colType||'').toUpperCase();
      if(t==='BOOL'||t==='BOOLEAN')return c.checked;
      if(t==='INTEGER')return parseInt(c.value)||0;
      if(t==='FLOAT')return parseFloat(c.value)||0;
      return c.value;
    });
  });
  return{columns:cols.map(function(c){return c.name;}),rows:rows};
}
function scVal(el){
  var t=(el.dataset.fieldType||'').toUpperCase();
  if(t==='BOOL'||t==='BOOLEAN')return el.checked;
  if(t==='INTEGER')return parseInt(el.value)||0;
  if(t==='FLOAT')return parseFloat(el.value)||0;
  return el.value;
}
function submit(){
  var adhoc=[];
  document.querySelectorAll('[data-scope="adhoc"]').forEach(function(el){adhoc.push({name:el.dataset.fieldName,value:scVal(el)});});
  document.querySelectorAll('.dt[data-table-scope="adhoc"]').forEach(function(t){adhoc.push({name:t.dataset.tableName,value:tblVal(t)});});
  var vlmap={};
  document.querySelectorAll('[data-scope="vl"]').forEach(function(el){var v=el.dataset.vlName;if(!vlmap[v])vlmap[v]=[];vlmap[v].push({name:el.dataset.fieldName,value:scVal(el)});});
  document.querySelectorAll('.dt[data-table-scope="vl"]').forEach(function(t){var v=t.dataset.vlName;if(!vlmap[v])vlmap[v]=[];vlmap[v].push({name:t.dataset.tableName,value:tblVal(t)});});
  var vld=Object.keys(vlmap).map(function(k){return{variableListName:k,variableInputs:vlmap[k]};});
  var ms=[];
  document.querySelectorAll('[data-group-id]').forEach(function(cb){ms.push({id:cb.dataset.groupId,name:cb.dataset.groupName,contentType:cb.dataset.contentType||'Group',isInclude:cb.checked,orderIndex:parseInt(cb.dataset.orderIndex)||0});});
  // Multiple documents can be attached to the same external-content slot: group checkboxes by
  // slot id, then emit one manualSelectContentItem PER CHECKED candidate (all sharing that
  // slot's id/name), or a single isInclude:false item if none are checked.
  var extGroups={};
  document.querySelectorAll('[data-external-candidate]').forEach(function(cb){
    var id=cb.dataset.externalCandidate;
    if(!extGroups[id])extGroups[id]={name:cb.dataset.externalName,orderIndex:parseInt(cb.dataset.orderIndex)||0,checked:[]};
    if(cb.checked)extGroups[id].checked.push(JSON.parse(cb.value));
  });
  Object.keys(extGroups).forEach(function(id){
    var g=extGroups[id];
    if(g.checked.length){
      g.checked.forEach(function(chosen){
        var item={id:id,name:g.name,isInclude:true,orderIndex:g.orderIndex,versionId:chosen.versionId,contentType:(chosen.format||'').toUpperCase()==='PDF'?'ResourcePDF':'LiveSlide'};
        if(chosen.sourceBlobId)item.sourceBlobId=chosen.sourceBlobId;
        ms.push(item);
      });
    }else{
      ms.push({id:id,name:g.name,contentType:'LiveSlide',isInclude:false,orderIndex:g.orderIndex});
    }
  });
  var p={teamSiteId:tsid,libraryContentVersionId:vid,adHocInputs:adhoc,outputs:sel};
  if(vld.length)p.variableListData=vld;
  if(ms.length)p.manualSelectContentInput={manualSelectContentItems:ms};
  var msg=JSON.stringify(p);
  var btn=document.getElementById('sub-btn');
  btn.disabled=true;btn.textContent='Submitting…';
  fetch('http://127.0.0.1:3099/submit/'+formToken,{method:'POST',headers:{'Content-Type':'application/json'},body:msg})
  .then(function(r){
    if(!r.ok)throw new Error('http '+r.status);
    btn.style.display='none';
    document.getElementById('done-msg').style.display='block';
  })
  .catch(function(){
    btn.disabled=false;btn.textContent='► Submit generation';
    var pt=document.getElementById('payload-text');
    pt.value=msg;
    document.getElementById('payload-box').style.display='block';
    btn.style.display='none';
    pt.select();
    try{navigator.clipboard.writeText(msg);}catch(e){}
    window.scrollTo(0,document.body.scrollHeight);
  });
}
function copyPayload(){
  var t=document.getElementById('payload-text');
  t.select();
  try{navigator.clipboard.writeText(t.value).then(function(){document.getElementById('copy-btn').textContent='Copied!';}).catch(function(){document.execCommand('copy');document.getElementById('copy-btn').textContent='Copied!';});}
  catch(e){try{document.execCommand('copy');document.getElementById('copy-btn').textContent='Copied!';}catch(e2){}}
}`;

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${css}</style></head><body>
<div class="title">${esc(templateName)} — inputs</div>
${scalarGrid}${tableHtml}${vlHtml}${msHtml}
${formSelHtml}
<div class="fmt-row"><div class="sl" style="margin-bottom:10px">Output format</div><div id="fmt-btns">${fmtBtnsHtml}</div></div>
<button id="sub-btn" class="sub" onclick="submit()">&#9654; Submit generation</button>
<div id="done-msg" style="display:none;margin-top:16px;padding:14px;background:#f0faf0;border:1.5px solid #b2dfb2;border-radius:8px;color:#2e7d32;font-weight:600;font-size:14px">&#10003; Submitted! Generation starting…</div>
<div id="payload-box" style="display:none;margin-top:16px;padding:14px;background:#f0f7ff;border:1.5px solid #90b8e8;border-radius:8px">
  <div style="font-weight:600;font-size:13px;margin-bottom:8px;color:#0055aa">Copy this payload and paste it into the chat:</div>
  <textarea id="payload-text" readonly style="width:100%;height:72px;font-size:11px;font-family:monospace;border:1px solid #b0c8e8;border-radius:4px;padding:6px;box-sizing:border-box;resize:none;background:#fff"></textarea>
  <button id="copy-btn" onclick="copyPayload()" style="margin-top:8px;padding:7px 20px;background:#0066cc;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600">&#128203; Copy to clipboard</button>
</div>
<script>${js}<\/script>
</body></html>`;
}

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
    sourceBlobId: gf(d, "sourceBlobId"),
    title: gf(d, "title"),
    format: gf(d, "format"),
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
        error: "manualSelectContentInput has items missing a resolved versionId.",
        detail: `Item(s) [${unresolved.map((i) => `"${i.name ?? i.id}"`).join(", ")}] have contentType requiring real content but no versionId. Call get_livedoc_inputs again and use the resolved "candidates" it attaches to this item, or call search_livedoc_content(query: <item name>, contentType: <item contentType>) directly, then set versionId to the result's contentVersionId — never fabricate one. If no match, set isInclude:false instead.`,
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
  const generatedLivedocId =
    (body.generatedLivedocId ?? body.id ?? body.GeneratedLivedocId ?? body.Id) as string | undefined;
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
  const result = await apiFetch(
    `/v3/generatedLivedocs/${args.generatedLivedocId}/outputs/${args.outputId}/content?redirect=false`
  );
  if (result.status !== 200) {
    return { error: `Download URL fetch failed (HTTP ${result.status})`, detail: result.body };
  }
  return result.body;
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
      fileName = match.fileName || `${fileName}.${(match.format ?? "").toLowerCase()}`;
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

async function handleWaitForFormSubmit(args: { token: string }): Promise<unknown> {
  const TIMEOUT_MS = 10 * 60 * 1000;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingForms.delete(args.token);
      reject(new Error("Form submission timed out after 10 minutes."));
    }, TIMEOUT_MS);
    pendingForms.set(args.token, (payload) => {
      clearTimeout(timer);
      resolve(payload);
    });
  });
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
  return {
    url,
    token,
    message: `Open this URL in your browser: ${url}\n\nAfter the form tab closes automatically, call get_form_result with token="${token}" to retrieve the generation result into this context.`,
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

const server = new Server(
  { name: "seismic-livedoc", version: "1.0.0" },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  let result: unknown;
  try {
    switch (name) {
      case "search_livedoc_templates":
        result = await handleSearchTemplates(args as { searchText?: string; page_size?: number });
        break;
      case "search_livedoc_content":
        result = await handleSearchContent(args as { query: string; contentType?: string; page_size?: number });
        break;
      case "get_livedoc_inputs": {
        const ir = await handleGetInputs(args as { teamSiteId: string; libraryContentVersionId: string });
        if ("error" in ir) {
          return { content: [{ type: "text" as const, text: JSON.stringify(ir, null, 2) }], isError: true };
        }
        if (ir.hasImageUpload) {
          return {
            content: [{ type: "text" as const, text: "This template requires image uploads. Use open_form_ui to open the full form in a browser tab." }],
          };
        }
        if (ir.isComplex) {
          return {
            content: [{ type: "text" as const, text: "This is a complex template. Use open_form_ui to open the full form in a browser tab instead of building an inline artifact." }],
          };
        }
        const formToken = generateToken();
        const formHtml = buildFormHtml(
          ir.templateName,
          ir.adhocInputs,
          ir.variableListData,
          ir.manualSelectContentInput,
          ir.forms,
          ir.teamSiteId,
          ir.libraryContentVersionId,
          formToken
        );
        return {
          content: [
            {
              type: "text" as const,
              text: [
                `Create an HTML artifact using EXACTLY the HTML below, verbatim, with no modifications — it already has a working, tested submit mechanism wired to this MCP server. Do NOT hand-write your own form HTML for this.`,
                `Immediately after creating the artifact, call wait_for_form_submit with token="${formToken}" — do not wait for the user to say anything first. That call blocks (up to 10 minutes) until the user clicks the form's submit button, then returns the exact payload to pass to submit_livedoc_generation.`,
                ``,
                "```html",
                formHtml,
                "```",
              ].join("\n"),
            },
          ],
        };
      }
      case "wait_for_form_submit":
        result = await handleWaitForFormSubmit(args as { token: string });
        break;
      case "submit_livedoc_generation":
        result = await handleSubmitGeneration(args as Parameters<typeof handleSubmitGeneration>[0]);
        break;
      case "get_generation_status":
        result = await handleGetStatus(args as { generatedLivedocId: string });
        break;
      case "open_form_ui":
        result = await handleOpenFormUi(args as { teamSiteId: string; libraryContentVersionId: string; context?: string });
        break;
      case "get_form_result":
        result = await handleGetFormResult(args as { token: string });
        break;
      case "login":
        result = await handleLogin(args as Parameters<typeof handleLogin>[0]);
        break;
      case "set_token":
        currentToken = (args as { token: string }).token;
        tokenIsManual = true;
        result = { ok: true, message: "Token updated. This token will not be auto-replaced by the credential-flow login on a 401 — call login explicitly to switch back to that flow." };
        break;
      case "get_generation_download_url": {
        const dlResult = await handleGetDownloadUrl(args as { generatedLivedocId: string; outputId: string });
        if (dlResult && typeof dlResult === "object" && "error" in (dlResult as object)) {
          return { content: [{ type: "text" as const, text: JSON.stringify(dlResult, null, 2) }], isError: true };
        }
        const dlBody = dlResult as Record<string, unknown>;
        const dlUrl = String(dlBody.url ?? dlBody.downloadUrl ?? dlBody.Url ?? dlBody.DownloadUrl ?? "");
        const dlFile = String(dlBody.fileName ?? dlBody.FileName ?? dlBody.name ?? dlBody.Name ?? "download");
        if (!dlUrl) {
          result = dlResult;
          break;
        }
        return {
          content: [{
            type: "text" as const,
            text: `✅ **${dlFile}** is ready.\n\nDownload link: [${dlFile}](${dlUrl})\n\n(Reproduce the markdown link above verbatim in your reply so the user can click it.)`,
          }],
        };
      }
      case "download_generation_output":
        result = await handleDownloadGenerationOutput(args as { generatedLivedocId: string; outputId: string; autoOpen?: boolean });
        break;
      // ── PPTX Auto-Tagging (PoC) ─────────────────────────────────────────────
      case "pptx_extract_shapes": {
        const a = (args ?? {}) as { pptxBase64: string; slideIndex?: number };
        const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
        const r = await fetch(`${pocBase}/api/pptx/extract`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            pptxBase64: a.pptxBase64,
            ...(a.slideIndex !== undefined ? { slideIndex: a.slideIndex } : {}),
          }),
        });
        if (!r.ok) throw new Error(`pptx_extract_shapes HTTP ${r.status}: ${await r.text()}`);
        result = await r.json();
        break;
      }

      case "pptx_auto_tag": {
        const a = (args ?? {}) as { pptxBase64: string; schema?: Record<string, string> };
        const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
        const r = await fetch(`${pocBase}/api/pptx/auto-tag`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ pptxBase64: a.pptxBase64, schema: a.schema ?? {} }),
        });
        if (!r.ok) throw new Error(`pptx_auto_tag HTTP ${r.status}: ${await r.text()}`);
        result = await r.json();
        break;
      }

      case "pptx_mark_shapes": {
        const a = (args ?? {}) as { pptxBase64: string; marks: unknown[] };
        const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
        const r = await fetch(`${pocBase}/api/pptx/mark`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ pptxBase64: a.pptxBase64, marks: a.marks }),
        });
        if (!r.ok) throw new Error(`pptx_mark_shapes HTTP ${r.status}: ${await r.text()}`);
        result = await r.json();
        // Summarize — omit the large base64 from output
        const { bindings, markedAt } = result as { bindings: unknown[]; markedAt: string };
        return {
          content: [{
            type: "text",
            text: JSON.stringify({ markedAt, bindingCount: bindings.length, bindings }, null, 2) +
              "\n\n(markedPptxBase64 available in full result — use pptx_mark_shapes result.pptxBase64 to save the file)",
          }],
        };
      }

      case "pptx_get_manifest": {
        const a = (args ?? {}) as { pptxBase64: string };
        const pocBase = process.env.POC_AUTOTAG_URL ?? "http://localhost:3001";
        const b64 = encodeURIComponent(a.pptxBase64);
        const r = await fetch(`${pocBase}/api/pptx/manifest?pptxBase64=${b64}`);
        if (!r.ok) throw new Error(`pptx_get_manifest HTTP ${r.status}: ${await r.text()}`);
        result = await r.json();
        break;
      }

      default:
        return {
          content: [{ type: "text", text: `Unknown tool: ${name}` }],
          isError: true,
        };
    }
  } catch (err) {
    return {
      content: [{ type: "text", text: `Tool error: ${err instanceof Error ? err.message : String(err)}` }],
      isError: true,
    };
  }

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
});

process.on("SIGTERM", () => process.exit(0));
process.on("SIGINT", () => process.exit(0));

// Auto-login on startup when credentials are available via env vars
if (!currentToken && DEFAULT_USERNAME && DEFAULT_PASSWORD) {
  await autoLogin();
}

const transport = new StdioServerTransport();
await server.connect(transport);
