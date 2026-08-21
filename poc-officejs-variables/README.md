# Livedoc Variables — Office.js POC

Tests three core Office.js capabilities as part of the VSTO → Office.js migration evaluation.

## What this POC tests

| # | Office.js API | Feature |
|---|---|---|
| ① | `customXmlParts.getByNamespaceAsync` | Load ad-hoc variables from document CustomXML |
| ② | `customXmlParts.addAsync` | Save variables to document CustomXML |
| ③ | `setSelectedDataAsync(token, CoercionType.Text)` | Insert `{{VarName}}` at cursor in a text box |
| ④ | `slide.shapes.addTable(rows, cols, opts)` | Insert table variable as a real PPT table |
| ⑤ | `presentation.getSelectedSlides()` | Get current slide for table insertion |

## Known gaps vs VSTO (tested here)

- `Shape.CustomerData` (per-shape CustomXML) — **not available** in Office.js. Variables are stored at doc level only.
- `Shape.Tags["key"]` — **not available**. No shape-level key-value store in Office.js.
- Chart object model — **not available** in Office.js for PowerPoint.

---

## Prerequisites

- Node.js 18+
- PowerPoint (Microsoft 365, desktop)
- Windows (for `office-addin-dev-certs`)

---

## Setup & Run

```bash
cd poc-officejs-variables

# 1. Install dependencies
npm install

# 2. Install trusted localhost SSL certs (one-time, may prompt for admin)
npm run install-certs

# 3. Start the local HTTPS server
npm start
```

The server starts at **https://localhost:3000** and auto-generates placeholder icons.

---

## Sideload into PowerPoint

1. Open PowerPoint (desktop)
2. **Insert** tab → **Get Add-ins** → **My Add-ins** tab → **Upload My Add-in**
3. Browse to `manifest.xml` in this folder → **Upload**
4. A **"Livedoc"** group appears in the **Home** ribbon tab
5. Click **"Variables"** to open the task pane

---

## Usage

### Chat Mode (local model)
1. Make sure Ollama is running locally and the model is available:
    - `ollama run qwen2.5:7b`
2. Start this add-in server:
    - `npm start`
3. Open the task pane and use the Chat Mode box.

Supported chat operations:
- Create scalar variable
- Create table variable
- Create computed variable
- Insert variable token or insert table shape
- Configure dynamic table binding on a selected shape
- Configure dynamic chart binding on a selected shape
- Configure dynamic image binding on a selected shape
- Run preview

Examples:
- Create a scalar variable ClientName with value Contoso and insert it.
- Create a table variable RevenueByRegion with columns Region, Amount and rows APAC 120, EMEA 90, AMER 150.
- Bind selected chart shape to RevenueByRegion using Region as labels and Amount as values.

Server environment options:
- `OLLAMA_MODEL` (default: `qwen2.5:7b`)
- `OLLAMA_CHAT_URL` (default: `http://localhost:11434/api/chat`)

### Variables tab
1. Enter a **Name** (e.g. `ClientName`), optional **Default Value** and **Group**
2. Click **+ Add Variable** — saved instantly to document CustomXML
3. On the slide, **double-click a text box** to enter edit mode (cursor visible)
4. Switch back to task pane, click **Insert** → `{{ClientName}}` appears at the cursor

### Table Variables tab
1. Enter a **Name** and paste **CSV** (first row = column headers)
2. Click **+ Add Table Variable** — saved to document CustomXML
3. Navigate to the slide where you want the table
4. Click **Insert Table** — a formatted PPT table is added to the current slide

### Groups
Variables can be organized into groups. Variables with the same **Group** value are displayed under a collapsible group header in the tree.

---

## Architecture note (Option A — sidecar service POC)

```
┌─────────────────────────────────────┐
│  PowerPoint Document (.pptx)        │
│  ┌───────────────────────────────┐  │
│  │  Document CustomXML Part      │  │  ← Office.js CAN read/write this
│  │  (all variable definitions)   │  │
│  └───────────────────────────────┘  │
│                                     │
│  Slide → Shape  (no shape-level     │  ← Office.js CANNOT attach metadata
│                  metadata in        │    to individual shapes
│                  Office.js)         │
└─────────────────────────────────────┘
         ↕  (future: REST calls)
┌─────────────────────────────────────┐
│  Sidecar Service                    │
│  Key: docId + shapePosition/name    │  ← Where per-shape metadata would live
│  Value: Livedoc PD XML              │    in a production Option A system
└─────────────────────────────────────┘
```

In production Option A, the doc-level CustomXML stores variable *definitions*. 
Per-shape metadata (Personalization Descriptors) would be stored in a REST service 
keyed by `documentId + shape identifier`, since `Shape.CustomerData` is unavailable.
