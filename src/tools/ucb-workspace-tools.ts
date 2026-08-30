import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  listWorkspaceSpaces,
  listWorkspaceFolders,
  submitUcbWorkspaceGeneration,
  getUcbWorkspaceGenerationStatus,
} from "../handlers/ucbWorkspace.js";
import { handleFindDocCenterProfile } from "../handlers/profile.js";

export function registerUcbWorkspaceTools(server: McpServer): void {
  server.registerTool(
    "find_doccenter_profile",
    {
      description:
        "Resolve a DocCenter profile's profileId and profileVersionId by profile name (and optionally teamSiteId). " +
        "Use this to fill in submit_ucb_workspace_generation's `origin.profileId`/`origin.profileVersionId` when the " +
        "user only gave you a profile name. If the target content is already published to a profile, " +
        "search_livedoc_templates/search_livedoc_content may return contentProfiles/profileVersionIds directly on " +
        "the matching result — check there first before calling this tool. " +
        "Does NOT resolve `origin.contentLocation`; ask the user for that.",
      inputSchema: {
        profileName: z.string().describe("Exact or partial DocCenter profile name to look up."),
        teamSiteId: z.string().optional().describe("Team site id to disambiguate when multiple teamsites have a profile with this name."),
      },
    },
    async (args) => {
      const result = await handleFindDocCenterProfile(args as { profileName: string; teamSiteId?: string });
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    }
  );

  server.registerTool(
    "list_workspace_spaces",
    {
      description:
        "List the Seismic Workspace spaces the current user can see. Call this first when the user wants to " +
        "generate a LiveDoc into Workspace, to get a spaceId for list_workspace_folders / submit_ucb_workspace_generation.",
      inputSchema: {},
    },
    async () => {
      const result = await listWorkspaceSpaces();
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    }
  );

  server.registerTool(
    "list_workspace_folders",
    {
      description:
        "List folders/files under a Workspace space. Omit folderId to list the space's root folders; " +
        "pass folderId to list the contents of that specific folder.",
      inputSchema: {
        spaceId: z.string().describe("Workspace space id, from list_workspace_spaces."),
        folderId: z.string().optional().describe("Folder id to drill into. Omit to list root folders."),
        offset: z.number().optional().describe("Pagination offset (default 0). Only used when folderId is set."),
        limit: z.number().optional().describe("Page size (default 100). Only used when folderId is set."),
      },
    },
    async (args) => {
      const result = await listWorkspaceFolders(
        args as { spaceId: string; folderId?: string; offset?: number; limit?: number }
      );
      return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
    }
  );

  server.registerTool(
    "submit_ucb_workspace_generation",
    {
      description:
        "Submit a LiveDoc generation whose output is written directly into a Seismic Workspace folder as a " +
        "linked file, instead of being downloaded. Requires exactly one output format. " +
        "IMPORTANT: the `origin` fields (profileId, profileVersionId, contentLocation) identify which DocCenter " +
        "profile/location this generation is being published against. Do NOT guess them: " +
        "if search_livedoc_templates/search_livedoc_content already returned contentProfiles/profileVersionIds for " +
        "the target content, use those; otherwise call find_doccenter_profile with the profile name to resolve " +
        "profileId/profileVersionId. `contentLocation` has no lookup — ask the user for it if not already provided. " +
        "Returns a generationId — call get_ucb_workspace_generation_status to poll; the Workspace file is " +
        "committed automatically once the generation is Ready.",
      inputSchema: {
        teamSiteId: z.string().describe("Team site identifier (UUID)."),
        libraryContentVersionId: z.string().describe("Content version identifier (UUID) of the LiveDoc template."),
        adHocInputs: z.array(z.object({ name: z.string(), value: z.any() })).describe("Array of {name, value} pairs for ALL ad hoc inputs."),
        outputs: z.array(z.object({
          format: z.string(),
          name: z.string().optional(),
          fileName: z.string().optional(),
        })).length(1).describe("Exactly one output format to generate."),
        variableListData: z.array(z.object({
          variableListName: z.string(),
          variableInputs: z.array(z.object({ name: z.string(), value: z.any() })),
        })).optional().describe("Variable list data from variableListDefinitions."),
        regionalFormat: z.string().optional().describe("Regional format culture name, e.g. \"en-US\"."),
        workspace: z.object({
          spaceId: z.string().describe("Destination Workspace space id, from list_workspace_spaces."),
          folderId: z.string().describe("Destination folder id, from list_workspace_folders."),
          name: z.string().describe("Name for the generated Workspace file (no extension needed)."),
          format: z.string().describe("File format, e.g. \"PPTX\" or \"DOCX\" — must match outputs[0].format."),
        }).describe("Where in Workspace the generated file should be created."),
        origin: z.object({
          profileId: z.string().describe("DocCenter profile id this generation is published against. Get it from search results' contentProfiles/profileVersionIds when present, or from find_doccenter_profile by profile name."),
          profileVersionId: z.string().describe("DocCenter profile version id. Get it from search results' profileVersionIds when present, or from find_doccenter_profile."),
          contentLocation: z.string().describe("DocCenter content location path. No lookup available — must be supplied by the caller."),
        }).describe("DocCenter origin metadata for this generation."),
      },
    },
    async (args) => {
      const result = await submitUcbWorkspaceGeneration(
        args as Parameters<typeof submitUcbWorkspaceGeneration>[0]
      );
      const body = result as Record<string, unknown>;
      if (body.error) {
        return { content: [{ type: "text" as const, text: JSON.stringify(body, null, 2) }], isError: true };
      }
      const generationId = String(body.generationId ?? "");
      return {
        content: [{
          type: "text" as const,
          text: [
            `UCB Workspace generation started. generationId: ${generationId}`,
            `NEXT: call get_ucb_workspace_generation_status with generationId="${generationId}". ` +
              `Keep polling every few seconds; once status is Ready the file will be committed to Workspace automatically.`,
          ].join("\n"),
        }],
      };
    }
  );

  server.registerTool(
    "get_ucb_workspace_generation_status",
    {
      description:
        "Check the status of a UCB Workspace generation job. Once status is Ready, this automatically commits " +
        "the generated file into Workspace and reports workspaceCommitted=true — no separate commit tool needed.",
      inputSchema: {
        generationId: z.string().describe("The generationId returned by submit_ucb_workspace_generation."),
      },
    },
    async (args) => {
      const result = await getUcbWorkspaceGenerationStatus(args as { generationId: string });
      const body = result as Record<string, unknown>;
      if (body.error && !body.status) {
        return { content: [{ type: "text" as const, text: JSON.stringify(body, null, 2) }], isError: true };
      }

      let nextStep: string;
      if (body.commitError) {
        nextStep = `Generation succeeded but committing to Workspace failed: ${body.commitError}. ` +
          `NEXT: call get_ucb_workspace_generation_status again with generationId="${body.generationId}" to retry the commit.`;
      } else if (body.workspaceCommitted) {
        nextStep = body.workspaceUrl
          ? `Done. The generated file has been committed to Workspace. Open it here: ${body.workspaceUrl}`
          : "Done. The generated file has been committed to Workspace.";
      } else if (body.status === "Failure" || body.status === "Cancelled") {
        nextStep = `Generation ${body.status}. Stop polling.`;
      } else {
        nextStep = `Still generating (status=${body.status}). NEXT: call get_ucb_workspace_generation_status again with generationId="${body.generationId}" in a few seconds.`;
      }

      return {
        content: [{
          type: "text" as const,
          text: [JSON.stringify(body, null, 2), nextStep].join("\n\n"),
        }],
      };
    }
  );
}
