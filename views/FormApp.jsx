import { useState, useEffect, useRef } from "react";
import S from "./styles.js";
import { useAppConnection } from "./hooks/useAppConnection.js";
import { useFormState } from "./hooks/useFormState.js";
import { useGenerationPoll } from "./hooks/useGenerationPoll.js";
import { LoginForm } from "./components/LoginForm.jsx";
import { DownloadButton } from "./components/DownloadButton.jsx";
import { SlidePreview } from "./components/SlidePreview.jsx";
import { ScalarInput } from "./components/ScalarInput.jsx";
import { TableInput } from "./components/TableInput.jsx";
import { Lightbox } from "./components/Lightbox.jsx";

export function FormApp() {
  const [schema,         setSchema]         = useState(null);
  const [errMsg,         setErrMsg]         = useState(null);
  const [previews,       setPreviews]       = useState([]);
  const [previewsFetched,setPreviewsFetched]= useState(false);
  const [previewImg,     setPreviewImg]     = useState(null);
  const [manualOutputFmts, setManualOutputFmts] = useState(["PPTX"]);
  const [fmtIdx,         setFmtIdx]         = useState(0);
  const [comboIdx,       setComboIdx]       = useState(0);
  const [thumbs,         setThumbs]         = useState({});
  const [previewErr,     setPreviewErr]     = useState(null);
  const [errors,         setErrors]         = useState({ scalars: {}, vlScalars: {}, tables: {} });
  const [wizStep,        setWizStep]        = useState(0);

  // ── Hooks ──────────────────────────────────────────────────────────────────

  // 1. Connection (handleFormLoad is a function declaration — hoisted)
  const { phase, setPhase, loginError, loginBusy, appRef, tokenRef, handlePanelLogin } =
    useAppConnection({ onFormLoad: handleFormLoad, onError: setErrMsg });

  // 2. Form values (resets automatically when schema changes)
  const { scalars, setScalars, tables, setTables,
          vlScalars, setVlScalars, vlTables, setVlTables,
          grpInc, setGrpInc, extSel, setExtSel, buildPayload, validateAll } = useFormState(schema);

  // 3. Generation lifecycle (needs setPhase from step 1)
  const { result, pollMsg, handleSubmit, resetResult, resumeFromResult } =
    useGenerationPoll({ appRef, tokenRef, setPhase, setErrMsg });

  // Reset all display state when a new form loads. Returns true if an existing result was
  // found and resumed (e.g. after a chat refresh remounted the panel mid/post-generation) —
  // the caller should then skip its own transition to the blank "ready" input form.
  function handleFormLoad(sc) {
    setSchema(sc);
    setFmtIdx(0);
    setComboIdx(0);
    setManualOutputFmts(["PPTX"]);
    setErrMsg(null);
    resetResult();
    setPreviews([]);
    setPreviewsFetched(false);
    setPreviewImg(null);
    setThumbs({});
    setPreviewErr(null);
    setErrors({ scalars: {}, vlScalars: {}, tables: {} });
    setWizStep(0);

    if (sc.existingResult) {
      return resumeFromResult(sc.existingResult, (gid, dls) => fetchPreviews(appRef.current, gid, dls));
    }
    return false;
  }

  // Fetch thumbnails for external content candidates whenever schema changes
  useEffect(() => {
    if (!schema || !appRef.current) return;
    const extItems = schema.externalContent ?? [];
    if (extItems.length === 0 || !schema.teamSiteId) return;

    const preloaded = {};
    extItems.forEach(slot => (slot.candidates ?? []).forEach(c => {
      if (c.versionId && c.thumbnailUrl) preloaded[c.versionId] = c.thumbnailUrl;
    }));
    if (Object.keys(preloaded).length > 0) setThumbs(preloaded);

    const allCandidates = extItems.flatMap(slot =>
      (slot.candidates ?? [])
        .filter(c => c.contentId && c.versionId && !c.thumbnailUrl)
        .map(c => ({ contentId: c.contentId, versionId: c.versionId }))
    );
    if (allCandidates.length === 0) return;

    appRef.current.callServerTool({
      name: "get_candidate_thumbnails",
      arguments: { teamSiteId: schema.teamSiteId, candidates: allCandidates },
    }).then(res => {
      const map = res?.structuredContent?.thumbnailMap ?? {};
      setThumbs(prev => ({ ...prev, ...map }));
    }).catch(() => {});
  }, [schema]);

  // Poll for AI-suggested sample values pushed via prefill_livedoc_form_values.
  // Stops as soon as a prefill is found, or after ~45s if none ever arrives.
  const prefillFoundRef = useRef(false);
  useEffect(() => {
    if (phase !== "ready" || !schema) return;
    prefillFoundRef.current = false;
    let attempts = 0;
    const interval = setInterval(async () => {
      if (prefillFoundRef.current) return;
      if (++attempts > 30) { clearInterval(interval); return; }
      const app = appRef.current;
      if (!app) return;
      try {
        const res = await app.callServerTool({ name: "get_form_prefill", arguments: { token: tokenRef.current } });
        const prefill = res?.structuredContent?.prefill;
        if (!prefill) return;
        prefillFoundRef.current = true;
        clearInterval(interval);
        if (prefill.scalars) setScalars(p => ({ ...p, ...prefill.scalars }));
        if (prefill.tables) setTables(p => ({ ...p, ...prefill.tables }));
        if (prefill.variableLists) {
          const vs = {}, vt = {};
          Object.entries(prefill.variableLists).forEach(([vlName, vl]) => {
            Object.entries(vl.scalars ?? {}).forEach(([k, v]) => { vs[`${vlName}|${k}`] = v; });
            Object.entries(vl.tables ?? {}).forEach(([k, v]) => { vt[`${vlName}|${k}`] = v; });
          });
          if (Object.keys(vs).length) setVlScalars(p => ({ ...p, ...vs }));
          if (Object.keys(vt).length) setVlTables(p => ({ ...p, ...vt }));
        }
      } catch { /* ignore poll errors */ }
    }, 1500);
    return () => clearInterval(interval);
  }, [phase, schema]);

  // ── Preview images ─────────────────────────────────────────────────────────

  async function fetchPreviews(app, generatedLivedocId, downloads) {
    const previewable = (downloads ?? []).filter(d => ["pptx", "pdf"].includes((d.format ?? "").toLowerCase()));
    if (!previewable.length) { setPreviewsFetched(true); return; }
    let networkError = false;
    try {
      const settled = await Promise.allSettled(previewable.map(d =>
        app.callServerTool({
          name: "get_preview_images",
          arguments: { generatedLivedocId, outputId: d.format.toLowerCase() },
        }).then(res => {
          const sc = res?.structuredContent;
          if (sc?.httpStatus && sc.httpStatus !== 200) return { format: d.format.toUpperCase(), images: [], unavailable: true };
          const images = (sc?.images ?? [])
            .filter(img => typeof img.url === "string" && img.url.startsWith("https://"));
          return { format: d.format.toUpperCase(), images, unavailable: false };
        })
      ));
      const results = settled.map((s, i) => {
        if (s.status === "fulfilled") return s.value;
        networkError = true;
        return { format: previewable[i].format.toUpperCase(), images: [], unavailable: false };
      });
      const nonempty = results.filter(r => r.images.length > 0);
      if (nonempty.length > 0) {
        setPreviews(nonempty);
        appRef.current?.sendSizeChanged({ width: 520, height: 680 });
      } else if (networkError) {
        setPreviewErr("Preview could not be loaded — network error.");
      }
    } finally {
      setPreviewsFetched(true);
    }
  }

  // ── Phase renders ──────────────────────────────────────────────────────────

  if (phase === "connecting" || phase === "loading") {
    return (
      <div style={{ ...S.page, display: "flex", alignItems: "center", gap: 10, color: "#888", padding: 32 }}>
        <span style={{ ...S.spinner, borderColor: "#888", borderTopColor: "transparent" }} />
        {phase === "connecting" ? "Connecting…" : "Loading form…"}
      </div>
    );
  }

  if (phase === "idle") {
    return (
      <div style={{ ...S.page, display: "flex", alignItems: "center", gap: 10, color: "#888", padding: 32 }}>
        <span style={{ color: "#2ea04f", fontWeight: 700 }}>✓</span>
        Signed in — waiting for a request…
      </div>
    );
  }

  if (phase === "login") {
    return <LoginForm app={appRef.current} onSubmit={handlePanelLogin} busy={loginBusy} error={loginError} />;
  }

  if (phase === "error") {
    return <div style={S.page}><div style={S.errBox}><b>Error:</b> {errMsg}</div></div>;
  }

  if (phase === "done") {
    const { generatedLivedocId, downloads = [], downloadUrls = [] } = result;
    const app = appRef.current;
    const primaryPreview = previews[0] ?? null;

    return (
      <div style={S.page}>
        <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>

        {previewsFetched && primaryPreview ? (
          <SlidePreview
            images={primaryPreview.images}
            templateName={schema.templateName}
            downloads={downloads}
            downloadUrls={downloadUrls}
            app={app}
            onZoom={i => setPreviewImg({
              images: primaryPreview.images.map((img, n) => ({
                url: img.url,
                title: `${primaryPreview.format} — Slide ${n + 1} of ${primaryPreview.images.length}`,
              })),
              idx: i,
            })}
          />
        ) : (
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
            {!previewsFetched && downloads.some(d => ["pptx","pdf"].includes((d.format??"").toLowerCase())) && (
              <div style={{ marginTop: 10, color: "#888", fontSize: 12, display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ ...S.spinner, borderColor: "#ccc", borderTopColor: "#888" }} />
                Loading preview…
              </div>
            )}
            {previewsFetched && previews.length === 0 && !previewErr && (
              <div style={{ marginTop: 10, color: "#aaa", fontSize: 11 }}>No slide preview available for this template.</div>
            )}
            {previewErr && (
              <div style={{ marginTop: 10, color: "#c00", fontSize: 11 }}>{previewErr}</div>
            )}
          </div>
        )}

        {previewsFetched && previews.length > 1 && previews.slice(1).map(p => {
          const lbImgs = p.images.map((img, n) => ({ url: img.url, title: `${p.format} — Slide ${n + 1} of ${p.images.length}` }));
          return (
            <div key={p.format} style={{ marginTop: 14 }}>
              <div style={{ ...S.sl, marginBottom: 8 }}>{p.format} Preview — {p.images.length} slide{p.images.length !== 1 ? "s" : ""}</div>
              <div style={{ display: "flex", overflowX: "auto", gap: 8, paddingBottom: 6 }}>
                {p.images.map((img, i) => (
                  <div key={img.index}
                    onClick={() => setPreviewImg({ images: lbImgs, idx: i })}
                    style={{ flexShrink: 0, width: 120, height: 80, borderRadius: 5, overflow: "hidden",
                      cursor: "zoom-in", border: "1.5px solid #dde3ea", background: "#f5f5f5",
                      boxShadow: "0 1px 3px rgba(0,0,0,.1)" }}
                    title={`Slide ${img.index + 1}`}>
                    <img src={img.url} alt={`Slide ${img.index + 1}`} style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                  </div>
                ))}
              </div>
            </div>
          );
        })}

        {previewImg && (
          <Lightbox
            images={previewImg.images}
            idx={previewImg.idx}
            onClose={() => setPreviewImg(null)}
            onNavigate={i => setPreviewImg(p => ({ ...p, idx: i }))}
          />
        )}
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

  // ── Ready: render full form ────────────────────────────────────────────────

  const { adhocScalars = [], adhocTables = [], variableLists = [], slideGroups = [], externalContent = [], formOptions = [], pageGroups = [] } = schema;
  const multiForm = formOptions.length > 1;
  const activeFmtOpt = formOptions[fmtIdx] ?? {};
  const combos = activeFmtOpt.outputCombos ?? [];

  // Wizard mode — when a template has several distinct groups (input fields, one or more
  // variable lists, content selection), one long scrolling form gets unwieldy. Split into
  // pages with Next/Back instead.
  //
  // Prefer the template author's own page boundaries (pageGroups, from the form definition's
  // FormElements tree) when available — that's the real structure they designed. Each
  // variable list always gets its own page regardless (variable lists come from a different
  // endpoint than FormElements, so they can't be matched into pageGroups).
  const usePageGroups = pageGroups.length > 1;
  const pageFieldSets = usePageGroups ? pageGroups.map(pg => new Set(pg.fieldNames)) : [];
  const scalarInPage = (idx, name) => pageFieldSets[idx]?.has(name) ?? false;
  const tableInPage = (idx, t) => scalarInPage(idx, t.name) || t.columns.some(c => scalarInPage(idx, c.name));

  // A page can be dedicated entirely to a real (non-AD_HOC) variable list's fields — those
  // never appear in adhocScalars/adhocTables, so fieldNames alone can't find them. Match each
  // page's variableListNames against the fetched variableLists[] by name so that page renders
  // the variable list inline instead of coming up blank (see formDefinition.ts PageGroup).
  const vlIndexByName = new Map(variableLists.map((vl, i) => [vl.name.toLowerCase(), i]));
  const pageVlIndices = usePageGroups
    ? pageGroups.map(pg => (pg.variableListNames ?? [])
        .map(n => vlIndexByName.get(n.toLowerCase()))
        .filter(i => i !== undefined))
    : [];
  const coveredVlIndices = new Set(pageVlIndices.flat());

  const inputSections = usePageGroups
    ? pageGroups.map((pg, i) => ({ key: `page:${i}`, label: pg.label }))
    : ((adhocScalars.length > 0 || adhocTables.length > 0) ? [{ key: "inputs", label: "Inputs" }] : []);
  const sections = [
    ...inputSections,
    ...variableLists
      .map((vl, i) => ({ key: `vl:${i}`, label: vl.name, i }))
      .filter(s => !coveredVlIndices.has(s.i)),
    ...((slideGroups.length > 0 || externalContent.length > 0) ? [{ key: "content", label: "Content" }] : []),
    { key: "output", label: "Output" },
  ];
  // Trigger pagination on GROUP count (distinct sections a user has to scroll past), not raw
  // field count — a template with 3 small variable lists is just as cluttered as one with 8
  // scalar fields. Real template-authored pages always page regardless of count.
  const groupCount = sections.length - 1; // exclude the always-present "Output" step
  const useWizard = usePageGroups || groupCount > 2;
  const step = useWizard ? Math.min(wizStep, sections.length - 1) : 0;
  const currentKey = sections[step]?.key;
  const currentVlIdx = currentKey?.startsWith("vl:") ? Number(currentKey.split(":")[1]) : null;
  const currentPageIdx = currentKey?.startsWith("page:") ? Number(currentKey.split(":")[1]) : null;
  const vlVisibleOnCurrentPage = i => currentPageIdx !== null && (pageVlIndices[currentPageIdx] ?? []).includes(i);
  const showSection = key => !useWizard || currentKey === key;

  // Scalars/tables belonging to an ARBITRARY section key — used for whatever the current
  // step is (visibleScalars/visibleTables below) and to check OTHER sections when scanning
  // for the first erroring page on failed submit.
  function fieldsForKey(key) {
    if (key === "inputs") return { scalars: adhocScalars, tables: adhocTables };
    if (key.startsWith("page:")) {
      const idx = Number(key.split(":")[1]);
      return {
        scalars: adhocScalars.filter(f => scalarInPage(idx, f.name)),
        tables: adhocTables.filter(t => tableInPage(idx, t)),
      };
    }
    return { scalars: [], tables: [] };
  }

  // Which scalars/tables belong on the CURRENT step — page-scoped when using real pages,
  // otherwise the whole "inputs" bucket at once (matches the old heuristic behavior).
  const { scalars: visibleScalars, tables: visibleTables } = !useWizard
    ? { scalars: adhocScalars, tables: adhocTables }
    : fieldsForKey(currentKey ?? "");
  // Separate "show all" (wizard off) from "wizard on but this isn't a vl: step" (show none) —
  // both cases could naively collapse to currentVlIdx === null, which would show every
  // variable list on every non-vl step (e.g. the Inputs page) instead of none of them.
  const showAllVariableLists = !useWizard;

  function validatePage(key, v) {
    if (key.startsWith("vl:")) {
      const vl = variableLists[Number(key.split(":")[1])];
      if (!vl) return true;
      return vl.scalars.every(f => !v.vlScalars[`${vl.name}|${f.name}`])
        && vl.tables.every(t => !v.tables[`${vl.name}|${t.name}`]);
    }
    if (key === "inputs" || key.startsWith("page:")) {
      const { scalars, tables } = fieldsForKey(key);
      const scalarNames = new Set(scalars.map(f => f.name));
      const tableNames = new Set(tables.map(t => t.name));
      const noAdhocErrors = Object.keys(v.scalars).every(n => !scalarNames.has(n))
        && Object.keys(v.tables).every(n => !tableNames.has(n));
      if (!noAdhocErrors) return false;
      // Also check any variable list(s) rendered inline on this page (see pageVlIndices).
      if (key.startsWith("page:")) {
        const idx = Number(key.split(":")[1]);
        for (const vlI of pageVlIndices[idx] ?? []) {
          const vl = variableLists[vlI];
          if (!vl) continue;
          const vlOk = vl.scalars.every(f => !v.vlScalars[`${vl.name}|${f.name}`])
            && vl.tables.every(t => !v.tables[`${vl.name}|${t.name}`]);
          if (!vlOk) return false;
        }
      }
      return true;
    }
    return true;
  }

  function goNext() {
    const v = validateAll(schema);
    setErrors(v);
    if (!validatePage(currentKey, v)) return;
    setWizStep(s => Math.min(s + 1, sections.length - 1));
  }

  function trySubmit() {
    const v = validateAll(schema);
    setErrors(v);
    const hasErrors = Object.keys(v.scalars).length > 0 || Object.keys(v.vlScalars).length > 0 || Object.keys(v.tables).length > 0;
    if (hasErrors) {
      if (useWizard) {
        // Jump back to the first step (in document order) that actually has an error.
        const badIdx = sections.findIndex(s => !validatePage(s.key, v));
        if (badIdx >= 0) setWizStep(badIdx);
      }
      return;
    }
    handleSubmit(
      () => buildPayload({ schema, fmtIdx, comboIdx, manualOutputFmts }),
      (gid, dls) => fetchPreviews(appRef.current, gid, dls)
    );
  }

  return (
    <div style={S.page}>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      <div style={S.title}>{schema.templateName}</div>

      {useWizard && (
        <div style={S.wizDots}>
          {sections.map((s, i) => (
            <div key={s.key} title={s.label} style={{ ...S.wizDot, ...(i === step ? S.wizDotOn : {}) }} />
          ))}
        </div>
      )}

      {visibleScalars.length > 0 && (
        <div style={S.grid}>
          {visibleScalars.map(f => (
            <ScalarInput key={f.name} field={f} value={scalars[f.name]} error={errors.scalars[f.name]}
              onChange={v => setScalars(p => ({ ...p, [f.name]: v }))} />
          ))}
        </div>
      )}

      {visibleTables.map(t => (
        <TableInput key={t.name} table={t} rows={tables[t.name] ?? []} error={errors.tables[t.name]}
          onChange={rows => setTables(p => ({ ...p, [t.name]: rows }))} />
      ))}

      {variableLists.map((vl, i) => (showAllVariableLists || currentVlIdx === i || vlVisibleOnCurrentPage(i)) && (
        <div key={vl.name} style={S.section}>
          <div style={S.sl}>
            Variable list — {vl.name}
            {vl.dataSourceName && <span style={S.badge}>{vl.dataSourceName}</span>}
          </div>
          {vl.scalars.length > 0 && (
            <div style={S.grid}>
              {vl.scalars.map(f => (
                <ScalarInput key={f.name} field={f} value={vlScalars[`${vl.name}|${f.name}`]}
                  error={errors.vlScalars[`${vl.name}|${f.name}`]}
                  onChange={v => setVlScalars(p => ({ ...p, [`${vl.name}|${f.name}`]: v }))} />
              ))}
            </div>
          )}
          {vl.tables.map(t => (
            <TableInput key={t.name} table={t} rows={vlTables[`${vl.name}|${t.name}`] ?? []}
              error={errors.tables[`${vl.name}|${t.name}`]}
              onChange={rows => setVlTables(p => ({ ...p, [`${vl.name}|${t.name}`]: rows }))} />
          ))}
        </div>
      ))}

      {showSection("content") && slideGroups.length > 0 && (
        <div style={S.section}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <div style={S.sl}>Content selection</div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => setGrpInc(p => Object.fromEntries(Object.keys(p).map(k => [k, true])))}
                style={{ fontSize: 11, color: "#0066cc", background: "none", border: "none", cursor: "pointer", padding: "2px 4px" }}>
                Select all
              </button>
              <button onClick={() => setGrpInc(p => Object.fromEntries(Object.keys(p).map(k => [k, false])))}
                style={{ fontSize: 11, color: "#0066cc", background: "none", border: "none", cursor: "pointer", padding: "2px 4px" }}>
                Unselect all
              </button>
            </div>
          </div>
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
                      <button
                        title="Preview"
                        onClick={e => { e.stopPropagation(); setPreviewImg({ images: [{ url: g.thumbnailUrl, title: g.name }], idx: 0 }); }}
                        style={{ position: "absolute", bottom: 28, right: 4, background: "rgba(0,0,0,0.45)", border: "none", borderRadius: "50%", width: 22, height: 22, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "#fff", fontSize: 12, lineHeight: 1 }}
                      >
                        👁
                      </button>
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

      {showSection("content") && externalContent.length > 0 && (
        <div style={S.section}>
          <div style={S.sl}>External content</div>
          {externalContent.map(slot => (
            <div key={slot.id} style={S.extItem}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                <div style={S.fl}>{slot.name}</div>
                {slot.candidates.length > 0 && (
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={() => setExtSel(p => ({ ...p, [slot.id]: new Set(slot.candidates.map(c => c.versionId)) }))}
                      style={{ fontSize: 11, color: "#0066cc", background: "none", border: "none", cursor: "pointer", padding: "2px 4px" }}>
                      Select all
                    </button>
                    <button onClick={() => setExtSel(p => ({ ...p, [slot.id]: new Set() }))}
                      style={{ fontSize: 11, color: "#0066cc", background: "none", border: "none", cursor: "pointer", padding: "2px 4px" }}>
                      Unselect all
                    </button>
                  </div>
                )}
              </div>
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

      {showSection("output") && multiForm && (
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

      {showSection("output") && combos.length > 1 && (
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
      {showSection("output") && combos.length === 1 && combos[0].length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={S.sl}>Output format</div>
          <button style={{ ...S.pill, ...S.pillOn }}>
            {combos[0].map(o => o.format).join(" + ")}
          </button>
        </div>
      )}
      {showSection("output") && (combos.length === 0 || (combos.length === 1 && combos[0].length === 0)) && (
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

      {useWizard ? (
        <div style={S.wizNav}>
          <button onClick={() => setWizStep(s => Math.max(s - 1, 0))} disabled={step === 0}
            style={{ ...S.wizBtn, ...(step === 0 ? { opacity: 0.4, cursor: "default" } : {}) }}>
            ← Back
          </button>
          {step < sections.length - 1
            ? <button onClick={goNext} style={{ ...S.wizBtn, background: "#0066cc", color: "#fff", borderColor: "#0066cc" }}>Next →</button>
            : <button onClick={trySubmit} style={{ ...S.sub, marginTop: 0, width: "auto", padding: "8px 22px" }}>▶ Submit generation</button>
          }
        </div>
      ) : (
        <button onClick={trySubmit} style={S.sub}>
          ▶ Submit generation
        </button>
      )}

      {previewImg && (
        <Lightbox
          images={previewImg.images}
          idx={previewImg.idx}
          onClose={() => setPreviewImg(null)}
          onNavigate={i => setPreviewImg(p => ({ ...p, idx: i }))}
        />
      )}
    </div>
  );
}
