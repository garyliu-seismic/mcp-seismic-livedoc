import React, { useState, useEffect, useRef } from "react";
import { createRoot } from "react-dom/client";
import { App } from "@modelcontextprotocol/ext-apps";

// ── Helpers ───────────────────────────────────────────────────────────────────

function emptyRow(columns) {
  const row = {};
  columns.forEach(c => {
    const t = (c.colType || "TEXT").toUpperCase();
    row[c.name] = (t === "BOOL" || t === "BOOLEAN") ? false : "";
  });
  return row;
}

function coerce(value, type) {
  const t = (type || "STRING").toUpperCase();
  if (t === "INTEGER") return value === "" || value === undefined ? 0 : (parseInt(value) || 0);
  if (t === "FLOAT")   return value === "" || value === undefined ? 0 : (parseFloat(value) || 0);
  if (t === "BOOL" || t === "BOOLEAN") return !!value;
  return value ?? "";
}

function coerceRow(row, columns) {
  const out = {};
  columns.forEach(c => {
    const t = (c.colType || "TEXT").toUpperCase();
    out[c.name] = (t === "INTEGER") ? (parseInt(row[c.name]) || 0)
      : (t === "FLOAT") ? (parseFloat(row[c.name]) || 0)
      : (t === "BOOL" || t === "BOOLEAN") ? !!row[c.name]
      : (row[c.name] ?? "");
  });
  return out;
}

function tableValue(rows, columns) {
  return {
    columns: columns.map(c => c.name),
    rows: rows.map(r => columns.map(c => coerceRow(r, columns)[c.name])),
  };
}

// ── Styles ────────────────────────────────────────────────────────────────────

const S = {
  page:    { fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif', padding: "16px 22px", background: "#fff", fontSize: 14, color: "#1d1d1f", maxWidth: 720 },
  title:   { fontSize: 16, fontWeight: 700, marginBottom: 18, paddingBottom: 8, borderBottom: "2px solid #e89520" },
  sl:      { fontSize: 12, fontWeight: 700, color: "#555", textTransform: "uppercase", letterSpacing: "0.4px", marginBottom: 8 },
  section: { marginBottom: 16, padding: "12px 14px", background: "#f9f9f9", border: "1px solid #ebebeb", borderRadius: 7 },
  grid:    { display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px 14px", marginBottom: 12 },
  fw:      { display: "flex", flexDirection: "column", gap: 4 },
  fl:      { fontSize: 12, fontWeight: 600, color: "#555" },
  fi:      { width: "100%", padding: "6px 9px", border: "1px solid #d0d0d0", borderRadius: 6, fontSize: 13, outline: "none", background: "#fff", boxSizing: "border-box" },
  bool:    { display: "flex", alignItems: "center", gap: 7, padding: "4px 0", fontSize: 13, cursor: "pointer" },
  tblWrap: { marginBottom: 12 },
  table:   { width: "100%", borderCollapse: "collapse", fontSize: 12 },
  th:      { padding: "5px 8px", background: "#f5f5f5", border: "1px solid #e0e0e0", fontWeight: 600, textAlign: "left", fontSize: 11, whiteSpace: "nowrap" },
  td:      { padding: "2px 5px", border: "1px solid #f0f0f0" },
  ci:      { width: "100%", border: "none", background: "transparent", fontSize: 12, outline: "none", padding: "3px 4px" },
  addBtn:  { marginTop: 5, padding: "3px 11px", fontSize: 11, color: "#0066cc", background: "none", border: "1px dashed #88aacc", borderRadius: 5, cursor: "pointer" },
  delBtn:  { color: "#bbb", background: "none", border: "none", cursor: "pointer", fontSize: 15, padding: "0 4px", lineHeight: 1 },
  pill:    { padding: "6px 16px", borderRadius: 20, border: "none", background: "#f0f0f0", color: "#444", fontSize: 12, fontWeight: 500, cursor: "pointer", marginRight: 6, marginBottom: 6 },
  pillOn:  { background: "#0066cc", color: "#fff" },
  grpRow:  { display: "flex", alignItems: "center", gap: 8, padding: "4px 0", cursor: "pointer", fontSize: 13 },
  extItem: { border: "1px solid #e5ecf5", borderRadius: 6, padding: "8px 10px", marginBottom: 8 },
  extOpt:  { display: "flex", alignItems: "center", gap: 6, padding: "3px 0", fontSize: 12 },
  badge:   { display: "inline-block", padding: "1px 6px", borderRadius: 100, fontSize: 11, fontWeight: 600, background: "#e8f0fe", color: "#0066cc", marginLeft: 6 },
  sub:     { marginTop: 16, width: "100%", padding: "10px", background: "#0066cc", color: "#fff", border: "none", borderRadius: 7, fontSize: 14, fontWeight: 600, cursor: "pointer" },
  okBox:   { marginTop: 14, padding: 14, background: "#f0faf0", border: "1.5px solid #b2dfb2", borderRadius: 8, color: "#2e7d32" },
  errBox:  { marginTop: 14, padding: 14, background: "#fff0f0", border: "1.5px solid #ffb0b0", borderRadius: 8, color: "#c00", fontSize: 13 },
  dlBtn:   { display: "inline-block", marginTop: 10, marginRight: 8, padding: "7px 16px", background: "#1a6fb5", color: "#fff", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: "pointer", textDecoration: "none" },
  spinner: { display: "inline-block", width: 14, height: 14, border: "2px solid #fff", borderTopColor: "transparent", borderRadius: "50%", verticalAlign: "middle", marginRight: 6, animation: "spin .7s linear infinite" },
};

// ── Sub-components ────────────────────────────────────────────────────────────

function DownloadButton({ app, url, fileName, label }) {
  const [state, setState] = React.useState("idle"); // idle | saving | done | error
  const [msg,   setMsg]   = React.useState("");
  async function handleClick() {
    if (state === "saving") return;
    setState("saving");
    try {
      const res = await app.callServerTool({ name: "download_output_file", arguments: { url, fileName } });
      const sc = res?.structuredContent;
      if (sc?.error) { setState("error"); setMsg(sc.error); return; }
      setState("done");
      setMsg(sc?.localPath ?? "Saved");
    } catch (e) {
      setState("error");
      setMsg(String(e));
    }
  }
  return (
    <div style={{ marginBottom: 8 }}>
      <button
        onClick={handleClick}
        disabled={state === "saving"}
        style={{
          padding: "8px 18px", background: state === "done" ? "#2e7d32" : "#1a6fb5",
          color: "#fff", border: "none", borderRadius: 6, fontSize: 13,
          fontWeight: 600, cursor: state === "saving" ? "wait" : "pointer",
          opacity: state === "saving" ? 0.7 : 1,
        }}
      >
        {state === "saving" ? "Saving…" : state === "done" ? "✓ Saved" : label}
      </button>
      {msg && (
        <div style={{ fontSize: 11, marginTop: 3, color: state === "error" ? "#c00" : "#555", wordBreak: "break-all" }}>
          {state === "error" ? `Error: ${msg}` : msg}
        </div>
      )}
    </div>
  );
}

function ScalarInput({ field, value, onChange }) {
  const t = (field.type || "STRING").toUpperCase();
  if (t === "BOOL" || t === "BOOLEAN") {
    return (
      <label style={S.bool}>
        <input type="checkbox" checked={!!value} onChange={e => onChange(e.target.checked)} />
        <span style={S.fl}>{field.name}</span>
      </label>
    );
  }
  return (
    <div style={S.fw}>
      <label style={S.fl}>{field.name}</label>
      <input
        type={t === "DATE" ? "date" : (t === "INTEGER" || t === "FLOAT") ? "number" : "text"}
        step={t === "FLOAT" ? "any" : undefined}
        value={value ?? ""}
        onChange={e => onChange(e.target.value)}
        style={S.fi}
      />
    </div>
  );
}

function TableInput({ table, rows, onChange }) {
  const addRow = () => onChange([...rows, emptyRow(table.columns)]);
  const removeRow = i => onChange(rows.filter((_, idx) => idx !== i));
  const setCell = (ri, col, val) => onChange(rows.map((r, i) => i === ri ? { ...r, [col]: val } : r));

  return (
    <div style={S.tblWrap}>
      <div style={{ ...S.fl, marginBottom: 4 }}>{table.name}</div>
      <table style={S.table}>
        <thead>
          <tr>
            {table.columns.map(c => <th key={c.name} style={S.th}>{c.name}</th>)}
            <th style={{ ...S.th, width: 24 }} />
          </tr>
        </thead>
        <tbody>
          {rows.map((row, ri) => (
            <tr key={ri}>
              {table.columns.map(c => {
                const t = (c.colType || "TEXT").toUpperCase();
                const bool = t === "BOOL" || t === "BOOLEAN";
                return (
                  <td key={c.name} style={S.td}>
                    {bool
                      ? <input type="checkbox" checked={!!row[c.name]} onChange={e => setCell(ri, c.name, e.target.checked)} />
                      : <input
                          type={(t === "INTEGER" || t === "FLOAT") ? "number" : t === "DATE" ? "date" : "text"}
                          step={t === "FLOAT" ? "any" : undefined}
                          value={row[c.name] ?? ""}
                          onChange={e => setCell(ri, c.name, e.target.value)}
                          style={S.ci}
                        />
                    }
                  </td>
                );
              })}
              <td style={S.td}><button onClick={() => removeRow(ri)} style={S.delBtn}>×</button></td>
            </tr>
          ))}
        </tbody>
      </table>
      <button onClick={addRow} style={S.addBtn}>+ Add row</button>
    </div>
  );
}

// ── Main form component ───────────────────────────────────────────────────────

function FormApp() {
  const appRef   = useRef(null);
  const tokenRef = useRef(null);

  const [phase,    setPhase]    = useState("connecting"); // connecting|loading|ready|submitting|polling|done|error
  const [schema,   setSchema]   = useState(null);
  const [errMsg,   setErrMsg]   = useState(null);
  const [result,   setResult]   = useState(null);    // { generatedLivedocId, downloadUrls }
  const [pollMsg,  setPollMsg]  = useState("");

  // form values
  const [scalars,  setScalars]  = useState({});   // { fieldName: rawValue }
  const [tables,   setTables]   = useState({});   // { tableName: [rows] }
  const [vlScalars,setVlScalars]= useState({});   // { "vl|fieldName": rawValue }
  const [vlTables, setVlTables] = useState({});   // { "vl|tableName": [rows] }
  const [grpInc,   setGrpInc]   = useState({});   // { groupId: bool }
  const [extSel,   setExtSel]   = useState({});   // { slotId: Set<versionId> }
  const [thumbs,         setThumbs]         = useState({});   // { versionId: thumbnailUrl }
  const [manualOutputFmts, setManualOutputFmts] = useState(["PPTX"]); // fallback when API has no output defs
  const [fmtIdx,   setFmtIdx]   = useState(0);    // selected formOption index
  const [comboIdx, setComboIdx] = useState(0);    // selected outputCombo index within form

  useEffect(() => {
    const app = new App({ name: "livedoc-form", version: "1.0.0" }, {});
    appRef.current = app;

    app.ontoolresult = async (event) => {
      const formToken = event?.structuredContent?.formToken;
      if (!formToken) return;
      tokenRef.current = formToken;
      setPhase("loading");
      try {
        const res = await app.callServerTool({ name: "get_form_schema", arguments: { token: formToken } });
        const sc = res?.structuredContent;
        if (!sc || sc.error) throw new Error(sc?.error ?? "No schema returned");

        // Initialise form state from schema
        const initScalars = {};
        (sc.adhocScalars ?? []).forEach(f => { initScalars[f.name] = ""; });
        const initTables = {};
        (sc.adhocTables ?? []).forEach(t => { initTables[t.name] = []; });
        const initVlScalars = {};
        const initVlTables  = {};
        (sc.variableLists ?? []).forEach(vl => {
          (vl.scalars ?? []).forEach(f => { initVlScalars[`${vl.name}|${f.name}`] = ""; });
          (vl.tables  ?? []).forEach(t => { initVlTables[`${vl.name}|${t.name}`]  = []; });
        });
        const initGrp = {};
        (sc.slideGroups ?? []).forEach(g => { initGrp[g.id] = g.defaultInclude; });
        const initExt = {};
        (sc.externalContent ?? []).forEach(e => { initExt[e.id] = new Set(); });

        setSchema(sc);
        setScalars(initScalars);
        setTables(initTables);
        setVlScalars(initVlScalars);
        setVlTables(initVlTables);
        setGrpInc(initGrp);
        setExtSel(initExt);
        setFmtIdx(0);
        setComboIdx(0);
        setPhase("ready");
        app.sendSizeChanged({ width: 520, height: 900 });

        // Fetch thumbnails for all external content candidates (fire-and-forget)
        const extItems = sc.externalContent ?? [];
        if (extItems.length > 0 && sc.teamSiteId) {
          // Seed thumbs map with URLs already present in the schema (from search results)
          const preloaded = {};
          extItems.forEach(slot => (slot.candidates ?? []).forEach(c => {
            if (c.versionId && c.thumbnailUrl) preloaded[c.versionId] = c.thumbnailUrl;
          }));
          if (Object.keys(preloaded).length > 0) setThumbs(preloaded);

          // Only call /v3/slides for candidates that had no thumbnailUrl in the search result
          const allCandidates = extItems.flatMap(slot =>
            (slot.candidates ?? [])
              .filter(c => c.contentId && c.versionId && !c.thumbnailUrl)
              .map(c => ({ contentId: c.contentId, versionId: c.versionId }))
          );
          if (allCandidates.length > 0) {
            app.callServerTool({
              name: "get_candidate_thumbnails",
              arguments: { teamSiteId: sc.teamSiteId, candidates: allCandidates },
            }).then(res => {
              const map = res?.structuredContent?.thumbnailMap ?? {};
              setThumbs(prev => ({ ...prev, ...map }));
            }).catch(() => { /* thumbnails are optional — ignore errors */ });
          }
        }
      } catch (e) {
        setErrMsg(String(e));
        setPhase("error");
      }
    };

    app.connect()
      .then(() => { /* stay in connecting until ontoolresult fires */ })
      .catch(e => { setErrMsg(String(e)); setPhase("error"); });
  }, []);

  // Poll for a new generation request when idle/done so the panel auto-reloads
  // without needing Claude Desktop to remount the iframe.
  const phaseRef = React.useRef(phase);
  phaseRef.current = phase;
  React.useEffect(() => {
    const interval = setInterval(async () => {
      if (phaseRef.current !== "done" && phaseRef.current !== "connecting") return;
      const app = appRef.current;
      if (!app) return;
      try {
        const res = await app.callServerTool({
          name: "get_latest_token",
          arguments: { currentToken: tokenRef.current ?? "" },
        });
        const sc = res?.structuredContent;
        if (!sc?.isNew || !sc?.formToken) return;
        // New generation request detected — load fresh schema
        tokenRef.current = sc.formToken;
        setPhase("loading");
        setErrMsg(null);
        setResult(null);
        const schemaRes = await app.callServerTool({ name: "get_form_schema", arguments: { token: sc.formToken } });
        const newSc = schemaRes?.structuredContent;
        if (!newSc || newSc.error) { setPhase("error"); setErrMsg(newSc?.error ?? "Schema load failed"); return; }
        const initScalars = {}; (newSc.adhocScalars ?? []).forEach(f => { initScalars[f.name] = ""; });
        const initTables  = {}; (newSc.adhocTables  ?? []).forEach(t => { initTables[t.name]  = []; });
        const initVlSc = {}; const initVlTb = {};
        (newSc.variableLists ?? []).forEach(vl => {
          (vl.scalars ?? []).forEach(f => { initVlSc[`${vl.name}|${f.name}`] = ""; });
          (vl.tables  ?? []).forEach(t => { initVlTb[`${vl.name}|${t.name}`]  = []; });
        });
        const initGrp = {}; (newSc.slideGroups ?? []).forEach(g => { initGrp[g.id] = g.defaultInclude; });
        const initExt = {}; (newSc.externalContent ?? []).forEach(e => { initExt[e.id] = new Set(); });
        setSchema(newSc); setScalars(initScalars); setTables(initTables);
        setVlScalars(initVlSc); setVlTables(initVlTb); setGrpInc(initGrp); setExtSel(initExt);
        setThumbs({}); setFmtIdx(0); setComboIdx(0); setManualOutputFmts(["PPTX"]);
        setPhase("ready");
      } catch { /* ignore poll errors */ }
    }, 2500);
    return () => clearInterval(interval);
  }, []);

  // ── payload builder ─────────────────────────────────────────────────────────

  function buildPayload() {
    const adHocInputs = [];
    (schema.adhocScalars ?? []).forEach(f => {
      adHocInputs.push({ name: f.name, value: coerce(scalars[f.name], f.type) });
    });
    (schema.adhocTables ?? []).forEach(t => {
      adHocInputs.push({ name: t.name, value: tableValue(tables[t.name] ?? [], t.columns) });
    });

    const variableListData = (schema.variableLists ?? []).map(vl => {
      const variableInputs = [];
      (vl.scalars ?? []).forEach(f => {
        variableInputs.push({ name: f.name, value: coerce(vlScalars[`${vl.name}|${f.name}`], f.type) });
      });
      (vl.tables ?? []).forEach(t => {
        variableInputs.push({ name: t.name, value: tableValue(vlTables[`${vl.name}|${t.name}`] ?? [], t.columns) });
      });
      // Skip VLs where all inputs are empty/zero
      const hasData = variableInputs.some(inp => {
        const v = inp.value;
        if (v === null || v === undefined || v === "" || v === 0) return false;
        if (typeof v === "object" && Array.isArray(v.rows)) return v.rows.length > 0;
        return true;
      });
      return hasData ? { variableListName: vl.name, variableInputs } : null;
    }).filter(Boolean);

    const msItems = [];
    (schema.slideGroups ?? []).forEach(g => {
      msItems.push({ id: g.id, name: g.name, contentType: g.contentType, isInclude: !!grpInc[g.id], orderIndex: g.orderIndex });
    });
    (schema.externalContent ?? []).forEach(slot => {
      const selected = extSel[slot.id] ?? new Set();
      if (selected.size > 0) {
        slot.candidates.filter(c => selected.has(c.versionId)).forEach(chosen => {
          msItems.push({
            id: slot.id, name: slot.name,
            contentType: chosen.format?.toUpperCase() === "PDF" ? "ResourcePDF" : "LiveSlide",
            isInclude: true, orderIndex: slot.orderIndex,
            versionId: chosen.versionId,
            ...(chosen.sourceBlobId ? { sourceBlobId: chosen.sourceBlobId } : {}),
          });
        });
      } else {
        msItems.push({ id: slot.id, name: slot.name, contentType: "LiveSlide", isInclude: false, orderIndex: slot.orderIndex });
      }
    });

    const formOpt = (schema.formOptions ?? [])[fmtIdx] ?? {};
    let outputs = (formOpt.outputCombos ?? [[]])[comboIdx] ?? [];
    // If the API returned no output definitions, fall back to the user's manual selection
    if (outputs.length === 0) outputs = manualOutputFmts.map(f => ({ format: f }));

    const payload = { adHocInputs, outputs };
    if (variableListData.length) payload.variableListData = variableListData;
    if (msItems.length) payload.manualSelectContentInput = { manualSelectContentItems: msItems };
    return payload;
  }

  // ── submit & poll ────────────────────────────────────────────────────────────

  async function handleSubmit() {
    const app = appRef.current;
    setPhase("submitting");
    setErrMsg(null);
    try {
      const payload = buildPayload();
      const submitRes = await app.callServerTool({
        name: "submit_form",
        arguments: { token: tokenRef.current, payload: JSON.stringify(payload) },
      });
      const sc = submitRes?.structuredContent;
      if (sc?.error) throw new Error(sc.error + (sc.detail ? ` — ${JSON.stringify(sc.detail)}` : ""));
      const generatedLivedocId = sc?.generatedLivedocId;
      if (!generatedLivedocId) throw new Error("No generatedLivedocId in response");

      // Poll until done
      setPhase("polling");
      const pollResult = await pollUntilDone(app, generatedLivedocId);
      setResult({ generatedLivedocId, downloadUrls: pollResult.downloadUrls ?? [], downloads: pollResult.downloads ?? [] });
      setPhase("done");
      app.sendSizeChanged({ width: 520, height: 500 });
    } catch (e) {
      setErrMsg(String(e));
      setPhase("error");
    }
  }

  async function pollUntilDone(app, id) {
    const MAX = 80; // ~4 minutes at 3s interval
    for (let i = 0; i < MAX; i++) {
      await new Promise(r => setTimeout(r, 3000));
      setPollMsg(`Generating… (${(i + 1) * 3}s)`);
      const res = await app.callServerTool({ name: "poll_generation", arguments: { generatedLivedocId: id } });
      const sc = res?.structuredContent;
      const status = sc?.status ?? "Unknown";
      if (status === "Completed") return sc ?? {};
      if (status === "Failed") throw new Error("Generation failed");
    }
    throw new Error("Generation timed out after 4 minutes");
  }

  // ── render ───────────────────────────────────────────────────────────────────

  if (phase === "connecting" || phase === "loading") {
    return (
      <div style={{ ...S.page, display: "flex", alignItems: "center", gap: 10, color: "#888", padding: 32 }}>
        <span style={{ ...S.spinner, borderColor: "#888", borderTopColor: "transparent" }} />
        {phase === "connecting" ? "Connecting…" : "Loading form…"}
      </div>
    );
  }

  if (phase === "error") {
    return <div style={S.page}><div style={S.errBox}><b>Error:</b> {errMsg}</div></div>;
  }

  if (phase === "done") {
    const { generatedLivedocId, downloads = [], downloadUrls = [] } = result;
    const app = appRef.current;
    return (
      <div style={S.page}>
        <div style={S.title}>{schema.templateName}</div>
        <div style={S.okBox}>
          <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 10 }}>✓ Generation complete!</div>
          {downloads.length > 0
            ? downloads.map((dl, i) => (
                <DownloadButton key={i} app={app} url={dl.url} fileName={dl.fileName} label={`↓ Download ${dl.format.toUpperCase()}`} />
              ))
            : downloadUrls.length > 0
              ? downloadUrls.map((url, i) => (
                  <DownloadButton key={i} app={app} url={url} fileName={`output-${i + 1}.pptx`} label={`↓ Download ${i + 1}`} />
                ))
              : <div style={{ fontSize: 13, marginTop: 6, color: "#555" }}>
                  ID: <code style={{ fontSize: 12 }}>{generatedLivedocId}</code><br />
                  Ask Claude to download this document.
                </div>
          }
        </div>
      </div>
    );
  }

  if (phase === "submitting" || phase === "polling") {
    return (
      <div style={S.page}>
        <div style={S.title}>{schema?.templateName ?? "LiveDoc"}</div>
        <div style={{ padding: 24, display: "flex", alignItems: "center", gap: 10, color: "#666" }}>
          <span style={{ ...S.spinner, borderColor: "#0066cc", borderTopColor: "transparent" }} />
          {phase === "submitting" ? "Submitting…" : pollMsg || "Generating…"}
        </div>
      </div>
    );
  }

  // ── ready: render full form ──────────────────────────────────────────────────

  const { adhocScalars = [], adhocTables = [], variableLists = [], slideGroups = [], externalContent = [], formOptions = [] } = schema;
  const multiForm = formOptions.length > 1;
  const activeFmtOpt = formOptions[fmtIdx] ?? {};
  const combos = activeFmtOpt.outputCombos ?? [];

  return (
    <div style={S.page}>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      <div style={S.title}>{schema.templateName}</div>

      {/* ── Ad hoc scalars ── */}
      {adhocScalars.length > 0 && (
        <div style={S.grid}>
          {adhocScalars.map(f => (
            <ScalarInput key={f.name} field={f} value={scalars[f.name]}
              onChange={v => setScalars(p => ({ ...p, [f.name]: v }))} />
          ))}
        </div>
      )}

      {/* ── Ad hoc tables ── */}
      {adhocTables.map(t => (
        <TableInput key={t.name} table={t} rows={tables[t.name] ?? []}
          onChange={rows => setTables(p => ({ ...p, [t.name]: rows }))} />
      ))}

      {/* ── Variable lists ── */}
      {variableLists.map(vl => (
        <div key={vl.name} style={S.section}>
          <div style={S.sl}>
            Variable list — {vl.name}
            {vl.dataSourceName && <span style={S.badge}>{vl.dataSourceName}</span>}
          </div>
          {vl.scalars.length > 0 && (
            <div style={S.grid}>
              {vl.scalars.map(f => (
                <ScalarInput key={f.name} field={f} value={vlScalars[`${vl.name}|${f.name}`]}
                  onChange={v => setVlScalars(p => ({ ...p, [`${vl.name}|${f.name}`]: v }))} />
              ))}
            </div>
          )}
          {vl.tables.map(t => (
            <TableInput key={t.name} table={t} rows={vlTables[`${vl.name}|${t.name}`] ?? []}
              onChange={rows => setVlTables(p => ({ ...p, [`${vl.name}|${t.name}`]: rows }))} />
          ))}
        </div>
      ))}

      {/* ── Slide groups ── */}
      {slideGroups.length > 0 && (
        <div style={S.section}>
          <div style={S.sl}>Content selection</div>
          {slideGroups.some(g => g.thumbnailUrl)
            ? <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                {slideGroups.map(g => {
                  const on = !!grpInc[g.id];
                  return (
                    <div key={g.id} onClick={() => setGrpInc(p => ({ ...p, [g.id]: !p[g.id] }))}
                      style={{ width: 130, borderRadius: 8, border: `2px solid ${on ? "#0066cc" : "#e0e0e0"}`,
                        overflow: "hidden", cursor: "pointer", position: "relative", background: "#f8f8f8" }}>
                      <img src={g.thumbnailUrl} alt={g.name}
                        style={{ width: "100%", height: 90, objectFit: "cover", display: "block" }} />
                      <div style={{ padding: "5px 7px", fontSize: 11, fontWeight: 600, color: "#333", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{g.name}</div>
                      {on && <div style={{ position: "absolute", top: 4, right: 4, background: "#0066cc", color: "#fff", borderRadius: "50%", width: 18, height: 18, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11 }}>✓</div>}
                    </div>
                  );
                })}
              </div>
            : slideGroups.map(g => (
                <label key={g.id} style={S.grpRow}>
                  <input type="checkbox" checked={!!grpInc[g.id]}
                    onChange={e => setGrpInc(p => ({ ...p, [g.id]: e.target.checked }))} />
                  {g.name}
                </label>
              ))
          }
        </div>
      )}

      {/* ── External content ── */}
      {externalContent.length > 0 && (
        <div style={S.section}>
          <div style={S.sl}>External content</div>
          {externalContent.map(slot => (
            <div key={slot.id} style={S.extItem}>
              <div style={{ ...S.fl, marginBottom: 8 }}>{slot.name}</div>
              {slot.candidates.length === 0
                ? <div style={{ fontSize: 12, color: "#888" }}>No candidates found</div>
                : <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                    {slot.candidates.map(c => {
                      const sel = extSel[slot.id] ?? new Set();
                      const checked = sel.has(c.versionId);
                      const thumb = c.thumbnailUrl || thumbs[c.versionId];
                      return (
                        <div
                          key={c.versionId}
                          onClick={() => setExtSel(p => {
                            const next = new Set(p[slot.id] ?? []);
                            checked ? next.delete(c.versionId) : next.add(c.versionId);
                            return { ...p, [slot.id]: next };
                          })}
                          style={{
                            width: 130, cursor: "pointer", borderRadius: 7,
                            border: checked ? "2.5px solid #0066cc" : "1.5px solid #dde3ea",
                            background: checked ? "#e8f0fe" : "#fff",
                            overflow: "hidden", transition: "border-color .15s",
                            boxShadow: checked ? "0 0 0 3px #b3cdf7" : "0 1px 3px rgba(0,0,0,.08)",
                          }}
                        >
                          {/* thumbnail area */}
                          <div style={{
                            width: "100%", height: 80, background: "#f0f0f0",
                            display: "flex", alignItems: "center", justifyContent: "center",
                            overflow: "hidden", position: "relative",
                          }}>
                            {thumb
                              ? <img src={thumb} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                              : <div style={{ fontSize: 22, color: "#ccc" }}>🖼</div>
                            }
                            {checked && (
                              <div style={{
                                position: "absolute", top: 5, right: 5,
                                background: "#0066cc", borderRadius: "50%",
                                width: 20, height: 20, display: "flex",
                                alignItems: "center", justifyContent: "center",
                                color: "#fff", fontSize: 12, fontWeight: 700,
                              }}>✓</div>
                            )}
                          </div>
                          {/* title row */}
                          <div style={{ padding: "6px 8px" }}>
                            <div style={{
                              fontSize: 11, fontWeight: checked ? 700 : 400,
                              color: checked ? "#0055b3" : "#333",
                              whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                            }} title={c.title || c.versionId}>
                              {c.title || c.versionId}
                            </div>
                            <div style={{ fontSize: 10, color: "#888", marginTop: 2 }}>{c.format}</div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
              }
            </div>
          ))}
        </div>
      )}

      {/* ── Form selector (multiple form names) ── */}
      {multiForm && (
        <div style={{ marginBottom: 14 }}>
          <div style={S.sl}>Select form</div>
          <div>
            {formOptions.map((opt, i) => (
              <button key={opt.name} onClick={() => { setFmtIdx(i); setComboIdx(0); }}
                style={{ ...S.pill, ...(i === fmtIdx ? S.pillOn : {}) }}>
                {opt.name}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* ── Output format combos ── */}
      {combos.length > 1 && (
        <div style={{ marginBottom: 14 }}>
          <div style={S.sl}>Output format</div>
          <div>
            {combos.map((combo, i) => {
              const label = combo.map(o => o.format).join(" + ") || "Default";
              return (
                <button key={i} onClick={() => setComboIdx(i)}
                  style={{ ...S.pill, ...(i === comboIdx ? S.pillOn : {}) }}>
                  {label}
                </button>
              );
            })}
          </div>
        </div>
      )}
      {combos.length === 1 && combos[0].length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={S.sl}>Output format</div>
          <button style={{ ...S.pill, ...S.pillOn }}>
            {combos[0].map(o => o.format).join(" + ")}
          </button>
        </div>
      )}
      {(combos.length === 0 || (combos.length === 1 && combos[0].length === 0)) && (
        <div style={{ marginBottom: 14 }}>
          <div style={S.sl}>Output format</div>
          {["PPTX", "PDF", "PPTX + PDF"].map(opt => {
            const fmts = opt.split(" + ");
            const active = JSON.stringify(manualOutputFmts) === JSON.stringify(fmts);
            return (
              <button key={opt} onClick={() => setManualOutputFmts(fmts)}
                style={{ ...S.pill, ...(active ? S.pillOn : {}) }}>
                {opt}
              </button>
            );
          })}
        </div>
      )}

      <button onClick={handleSubmit} style={S.sub}>▶ Submit generation</button>
    </div>
  );
}

// ── Mount ─────────────────────────────────────────────────────────────────────

createRoot(document.getElementById("root")).render(<FormApp />);
