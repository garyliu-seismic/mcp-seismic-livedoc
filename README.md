# mcp-seismic-livedoc

Model Context Protocol (MCP) server for Seismic LiveDoc generation. Exposes LiveDoc templates, content search, and document generation as MCP tools consumable by Claude Desktop and other MCP-compatible agents.

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

## Configuration

The server is configured via environment variables:

| Variable | Description |
|---|---|
| `SEISMIC_BASE_URL` | Base URL for the LiveDoc API |
| `FORM_APP_URL` | URL of the form UI dev server (default: `http://localhost:5173`) |
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
  handlers/         # Business logic (content, generation, inputs, panel)
  ipc/              # Temp-file IPC between the server and the React panel
  tools/            # MCP tool registrations
  utils/            # Debug logging, OS utilities
views/              # Vite-built React form shell (form-shell.html)
dist/               # Compiled output (generated — not committed)
```
