import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { FORM_API_BASE } from "../config.js";

const POC_AUTOTAG_URL = process.env.POC_AUTOTAG_URL ?? FORM_API_BASE;

export function registerPptxAutotagTools(server: McpServer): void {
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
      const pocBase = process.env.POC_AUTOTAG_URL ?? POC_AUTOTAG_URL;
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
      const pocBase = process.env.POC_AUTOTAG_URL ?? POC_AUTOTAG_URL;
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
      const pocBase = process.env.POC_AUTOTAG_URL ?? POC_AUTOTAG_URL;
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
      const pocBase = process.env.POC_AUTOTAG_URL ?? POC_AUTOTAG_URL;
      const r = await fetch(`${pocBase}/api/pptx/manifest?pptxBase64=${encodeURIComponent(a.pptxBase64)}`);
      if (!r.ok) throw new Error(`pptx_get_manifest HTTP ${r.status}: ${await r.text()}`);
      return { content: [{ type: "text" as const, text: JSON.stringify(await r.json(), null, 2) }] };
    }
  );
}
