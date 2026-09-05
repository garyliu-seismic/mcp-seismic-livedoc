# mcp-seismic-livedoc

Model Context Protocol (MCP) server that lets Claude Desktop generate Seismic LiveDoc documents through a conversational interface. The agent searches templates, opens an interactive React form panel for user input, drives the generation job, and downloads the finished file — all without leaving the chat.

## How it works

```
Claude Desktop (chat)
    │
    │  stdio (MCP protocol)
    ▼
mcp-seismic-livedoc (Node.js process)
    │
    ├── Auth layer ──────────────────── Seismic Auth Service (OAuth2 browser-login)
    │    └─ auto-login on startup          token saved to disk, refreshed every 60 s
    │
    ├── Seismic LiveDoc API ─────────── SEISMIC_BASE_URL/v3/...
    │    ├─ search templates
    │    ├─ fetch template input schema
    │    ├─ submit generation job
    │    └─ poll status / get download URL
    │
    └── App Panel (React) ───────────── embedded HTML served as MCP resource
         └─ communicates back to server via tool calls over the same MCP connection
```

### End-to-end generation flow

1. **Search** — Claude calls `search_livedoc_templates` to find the template by name and get its `teamSiteId` / `libraryContentVersionId`.

2. **Open form** — Claude calls `get_livedoc_inputs`. The server:
   - Fetches the template's full input schema from the LiveDoc API (`/v3/teamsites/{id}/livedocVersions/{id}`).
   - Resolves any "manual-select content" slots by pre-fetching candidates from the Seismic library.
   - Generates a unique `formToken` and writes the schema to a temp file (`%TEMP%/mcp-livedoc-schema-{token}.json`).
   - Writes a "latest token" pointer that the already-open panel polls to detect new requests.
   - Opens the **App panel** (a Vite-built React single-file HTML served as an MCP resource) in Claude Desktop's side panel.

3. **User fills the form** — The React panel calls internal MCP tools (`get_form_schema`, `get_form_prefill`, `get_auth_status`) to load state and render fields. The user completes inputs and clicks **Submit**.

4. **Panel drives generation** — On submit the panel calls `submit_form` → the server posts to the LiveDoc generation API. The panel then polls `poll_generation` every few seconds until the job reaches `Completed` or `Failed`. Progress (slide previews, status) is shown inline.

5. **Claude retrieves the result** — When Claude detects generation is done it calls `get_panel_result` with the `formToken` to read the result from `%TEMP%/mcp-livedoc-result-{token}.json`, then calls `download_generation_output` to save the file to the user's Downloads folder and open it.

### Token / IPC model

The Node.js MCP server process and the React panel (running inside Claude Desktop's renderer) share state through **temp files** rather than in-memory, because they run in separate processes:

| File | Purpose |
|---|---|
| `mcp-livedoc-schema-{token}.json` | Form schema written by server, read by panel via `get_form_schema` |
| `mcp-livedoc-latest-token.json` | Latest `formToken` pointer; panel polls this to auto-reload on new request |
| `mcp-livedoc-result-{token}.json` | Generation result written by panel, read by Claude via `get_panel_result` |
| `mcp-livedoc-prefill-{token}.json` | Optional pre-filled values written by `prefill_livedoc_form_values` |
| `mcp-livedoc-token.json` | Persisted bearer token; survives MCP server restarts |
| `mcp-livedoc-crash.log` | Crash/unhandled-rejection log |

### Authentication

On startup the server attempts auto-login using `AUTH_USERNAME`/`AUTH_PASSWORD`/`AUTH_TENANT` (or a previously saved token from disk). A background timer re-runs auto-login when the JWT is within 5 minutes of expiry. The `apiFetch` client also retries on 401/403 by picking up a token saved by the panel's own login flow — so a login performed in the panel is immediately available to the server without a restart.

## How Claude should use this MCP

### Normal flow — form panel (recommended)

This is the intended path for all interactive generation. Claude must follow this exact sequence and must not deviate:

```
1. search_livedoc_templates(searchText)
      → returns { contentVersionId, teamSiteId, name }

2. get_livedoc_inputs(teamSiteId, libraryContentVersionId)
      → opens the App panel with the input form
      → returns formToken

3. [optional] prefill_livedoc_form_values(formToken, scalars, tables, variableLists)
      → fills suggested values into the open form for the user to review

4. [user fills the form and clicks Submit in the panel]
      → panel posts to the LiveDoc API and polls generation internally
      → panel writes result to %TEMP%/mcp-livedoc-result-{token}.json

5. get_panel_result(formToken)   ← call when user says "done" or "generation finished"
      → returns { generatedLivedocId, status, downloads, templateName }

6. download_generation_output(generatedLivedocId, outputId)
      → saves file to Downloads folder and opens it
```

**Rules Claude must follow:**

- **Always call `search_livedoc_templates` first** — never ask the user for `teamSiteId` or `libraryContentVersionId`; those are internal IDs resolved by the search.
- **Never ask for credentials in chat** — authentication runs silently via env vars or the panel's sign-in form. If auth fails, tell the user to sign in through the panel.
- **Never call `submit_form`, `get_form_schema`, or `submit_livedoc_generation` to bypass the visible form** — these skip the user's review step. The only exception is the headless flow described below.
- **`prefill_livedoc_form_values` is for suggestions only** — it writes values into the open form; the user still reviews and submits. It must not be confused with submitting a generation.
- **Only use field names that appear in the schema from `get_livedoc_inputs`** — passing an unknown name to `prefill_livedoc_form_values` returns an error listing the valid names.

### Headless flow — no panel

Use this only when the user explicitly provides all input values and does not want the interactive form (e.g. bulk generation, scripted automation):

```
1. search_livedoc_templates(searchText)

2. submit_livedoc_generation(teamSiteId, libraryContentVersionId, adHocInputs, outputs)
      → returns generatedLivedocId

3. get_generation_status(generatedLivedocId)   ← repeat every few seconds until allDone=true

4. download_generation_output(generatedLivedocId, outputId)
```

**When templates have `manualSelectContentInput` slots** (content that must be chosen from the Seismic library), always use the panel flow — the server pre-fetches candidates and surfaces them in the form. Never try to resolve `versionId` values yourself.

### Summarising results

When `get_panel_result` returns `status: "Completed"`, Claude should:
1. Render an HTML artifact — a card with a green check, the template name, a downloads list with per-format icons and clickable links, and an expiry note.
2. Call `download_generation_output` for each output to save files locally.

### Handling auth failures

| Symptom | What to do |
|---|---|
| Tool returns 401/403 | Auto-retry runs automatically; if it keeps failing, tell the user to sign in via the panel |
| Panel shows the login screen | User fills tenant/username/password in the panel; the saved token is picked up by the server on the next tool call without any restart |
| `SEISMIC_API_TOKEN` set in env | Token is used as-is; auto-login is skipped; use `set_token` to rotate it |

### Common tool errors

These are the errors tools actually return for invalid input — recognize them and react as described instead of retrying blindly or asking the user for internal IDs.

| Tool | Trigger | Response shape |
|---|---|---|
| `prefill_livedoc_form_values` | Field/table/variable-list name not in the schema from `get_livedoc_inputs` | `isError: true`, text: `error: unknown field name(s): scalars.Bogus. Valid names — scalars: CompanyName, Amount; tables: LineItems; variableLists: Signers. Re-call prefill_livedoc_form_values using only these exact names.` — re-call using only the listed names, never guess a fix |
| `prefill_livedoc_form_values` | `formToken` doesn't match an open form (panel was closed, or a stale/typo'd token) | `isError: true`, text: `error: no form open for this formToken (schema not found)` — tell the user to reopen the form via `get_livedoc_inputs`; do not retry the same token |
| `get_panel_result` | Called before the user has submitted the form in the panel | Not an error — text: `No result yet — generation has not started or the form has not been submitted. Check the App panel.` — wait for the user to confirm, then call again |
| `submit_livedoc_generation` / any authenticated call | Token missing or expired and auto-login fails | HTTP 401/403 surfaced in the tool response — do not ask for the password in chat; tell the user to sign in via the panel's login form |
| `search_livedoc_templates` | No template matches the search text | Empty result list, not an error | Ask the user to refine the search text; don't fabricate a `teamSiteId`/`libraryContentVersionId` |

### What Claude should never do

- Ask the user for `teamSiteId`, `libraryContentVersionId`, `contentVersionId`, or `generatedLivedocId` — always resolve these via tools.
- Ask the user for their password in chat — credentials must only flow through env vars or the panel's login form.
- Call `get_form_schema`, `submit_form`, or `poll_generation` directly — those are internal tools used by the React panel only (`visibility: ["app"]`).
- Fabricate `versionId` values for manual-select content slots.

---

## Prerequisites

- Node.js 20+
- npm 9+

## Installation

```bash
npm install
```

## Build

```bash
# Build TypeScript only
npm run build

# Build the React form shell (Vite) + TypeScript
npm run build:all
```

## Development

```bash
# Run directly with tsx (no compile step)
npm run dev
```

## Testing

```bash
npm test
```

Runs Node's built-in test runner (via `tsx`) over `src/**/*.test.ts`. Coverage today focuses on the temp-file IPC layer (`src/ipc/temp-file.ts`) and the `prefill_livedoc_form_values` field-name validation (`src/handlers/prefillValidation.ts`).

## Configuration

The server is configured via environment variables:

| Variable | Description |
|---|---|
| `SEISMIC_BASE_URL` | Base URL for the LiveDoc API |
| `AUTH_SERVICE_URI` | Seismic auth service base URL |
| `AUTH_TENANT` | Tenant slug used for auto-login |
| `AUTH_CLIENT_ID` | OAuth client ID for auto-login |
| `AUTH_CLIENT_SECRET` | OAuth client secret for auto-login |
| `AUTH_USERNAME` | Username for auto-login |
| `AUTH_PASSWORD` | Password for auto-login |
| `SEISMIC_API_TOKEN` | Bearer token (overrides auto-login when set) |
| `POC_AUTOTAG_URL` | Base URL for the PoC PPTX autotag server |

A ready-to-use `.mcp.json` file is included at the repo root for Claude Desktop registration (update credentials before use).

## Claude Desktop Setup

Add the following to your `claude_desktop_config.json` (typically `%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "seismic-livedoc": {
      "command": "node",
      "args": ["<absolute-path-to-repo>/dist/index.js"],
      "env": {
        "SEISMIC_BASE_URL": "https://...",
        "AUTH_SERVICE_URI": "https://...",
        "AUTH_TENANT": "...",
        "AUTH_CLIENT_ID": "...",
        "AUTH_CLIENT_SECRET": "...",
        "AUTH_USERNAME": "...",
        "AUTH_PASSWORD": "..."
      }
    }
  }
}
```

## Available Tools

### Authentication
| Tool | Description |
|---|---|
| `login` | Authenticate with username/password and store the token |
| `set_token` | Manually set a bearer token |
| `check_auth` | Verify the current token is valid |
| `get_auth_status` | Return token expiry and validity details |
| `get_auth_config` | Show the resolved auth configuration |
| `panel_login` | Trigger login from within the App panel |
| `get_latest_token` | Return the current bearer token (for panel use) |

### Content & Templates
| Tool | Description |
|---|---|
| `search_livedoc_templates` | Search available LiveDoc templates |
| `search_livedoc_content` | Search Seismic content library |

### Generation
| Tool | Description |
|---|---|
| `prefill_livedoc_form_values` | Pre-populate form fields before opening the panel |
| `get_panel_result` | Retrieve the user's form submission from the panel |
| `submit_livedoc_generation` | Submit a generation job directly (bypasses panel) |
| `get_generation_status` | Poll the status of a running generation job |
| `get_generation_download_url` | Get a signed download URL for a completed job |
| `download_generation_output` | Download the generated file to the local machine |

### UCB Workspace Generation
| Tool | Description |
|---|---|
| `list_workspace_spaces` | List the Workspace spaces the current user can see |
| `list_workspace_folders` | List root folders in a space, or items in a specific folder |
| `submit_ucb_workspace_generation` | Submit a generation whose output is written directly into a Workspace folder |
| `get_ucb_workspace_generation_status` | Poll generation status; auto-commits the file to Workspace once Ready |

### Panel (App UI)
| Tool | Description |
|---|---|
| `get_form_schema` | Return the JSON schema for a template's input form |
| `get_form_prefill` | Return pre-filled values for the form |
| `submit_form` | Submit form values from within the App panel |
| `poll_generation` | Poll generation progress from within the App panel |
| `download_output_file` | Download the output file from within the App panel |
| `get_preview_images` | Return slide preview images for a completed job |
| `get_candidate_thumbnails` | Return template thumbnail images |

### PPTX Autotag (PoC)
| Tool | Description |
|---|---|
| `pptx_extract_shapes` | Extract all shapes from a PPTX file |
| `pptx_auto_tag` | Automatically tag shapes with LiveDoc field names |
| `pptx_mark_shapes` | Write tags back into the PPTX file |
| `pptx_get_manifest` | Return the LiveDoc manifest for a tagged PPTX |

### Debug
| Tool | Description |
|---|---|
| `debug_environment` | Dump resolved environment variables and auth state |
| `log_debug_message` | Write a message to the MCP debug log |

## Project Structure

```
src/
  index.ts          # Entry point — wires server, tools, and token lifecycle
  config.ts         # Reads Claude Desktop config for cowork file paths
  types.ts          # Shared TypeScript types
  api/              # Seismic API client (apiFetch)
  auth/             # Token state, auto-login, JWT utilities
  handlers/         # Business logic (content, generation, inputs, panel, ucbWorkspace)
  ipc/              # Temp-file IPC between the server and the React panel
  tools/            # MCP tool registrations
  utils/            # Debug logging, OS utilities
views/              # Vite-built React form shell (form-shell.html)
dist/               # Compiled output (generated — not committed)
```
