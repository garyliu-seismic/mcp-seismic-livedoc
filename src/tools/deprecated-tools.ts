import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { handleOpenFormUi, handleGetFormResult } from "../handlers/panel.js";

export function registerDeprecatedTools(server: McpServer): void {
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
}
