import { useState, useEffect } from "react";
import { coerce, tableValue } from "../utils/formUtils.js";

export function useFormState(schema) {
  const [scalars,   setScalars]   = useState({});
  const [tables,    setTables]    = useState({});
  const [vlScalars, setVlScalars] = useState({});
  const [vlTables,  setVlTables]  = useState({});
  const [grpInc,    setGrpInc]    = useState({});
  const [extSel,    setExtSel]    = useState({});

  useEffect(() => {
    if (!schema) return;
    const initScalars = {};
    (schema.adhocScalars ?? []).forEach(f => { initScalars[f.name] = ""; });
    const initTables = {};
    (schema.adhocTables ?? []).forEach(t => { initTables[t.name] = []; });
    const initVlScalars = {};
    const initVlTables  = {};
    (schema.variableLists ?? []).forEach(vl => {
      (vl.scalars ?? []).forEach(f => { initVlScalars[`${vl.name}|${f.name}`] = ""; });
      (vl.tables  ?? []).forEach(t => { initVlTables[`${vl.name}|${t.name}`]  = []; });
    });
    const initGrp = {};
    (schema.slideGroups ?? []).forEach(g => { initGrp[g.id] = g.defaultInclude; });
    const initExt = {};
    (schema.externalContent ?? []).forEach(e => { initExt[e.id] = new Set(); });

    setScalars(initScalars);
    setTables(initTables);
    setVlScalars(initVlScalars);
    setVlTables(initVlTables);
    setGrpInc(initGrp);
    setExtSel(initExt);
  }, [schema]);

  function buildPayload({ schema: sc, fmtIdx, comboIdx, manualOutputFmts }) {
    const adHocInputs = [];
    (sc.adhocScalars ?? []).forEach(f => {
      adHocInputs.push({ name: f.name, value: coerce(scalars[f.name], f.type) });
    });
    (sc.adhocTables ?? []).forEach(t => {
      adHocInputs.push({ name: t.name, value: tableValue(tables[t.name] ?? [], t.columns) });
    });

    const variableListData = (sc.variableLists ?? []).map(vl => {
      const variableInputs = [];
      (vl.scalars ?? []).forEach(f => {
        variableInputs.push({ name: f.name, value: coerce(vlScalars[`${vl.name}|${f.name}`], f.type) });
      });
      (vl.tables ?? []).forEach(t => {
        variableInputs.push({ name: t.name, value: tableValue(vlTables[`${vl.name}|${t.name}`] ?? [], t.columns) });
      });
      const hasData = variableInputs.some(inp => {
        const v = inp.value;
        if (v === null || v === undefined || v === "" || v === 0 || v === false) return false;
        if (typeof v === "object" && Array.isArray(v.rows)) return v.rows.length > 0;
        return true;
      });
      return hasData ? { variableListName: vl.name, variableInputs } : null;
    }).filter(Boolean);

    const msItems = [];
    (sc.slideGroups ?? []).forEach(g => {
      msItems.push({ id: g.id, name: g.name, contentType: g.contentType, isInclude: !!grpInc[g.id], orderIndex: g.orderIndex });
    });
    (sc.externalContent ?? []).forEach(slot => {
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

    const formOpt = (sc.formOptions ?? [])[fmtIdx] ?? {};
    let outputs = (formOpt.outputCombos ?? [[]])[comboIdx] ?? [];
    if (outputs.length === 0) outputs = manualOutputFmts.map(f => ({ format: f }));

    const payload = { adHocInputs, outputs };
    if (variableListData.length) payload.variableListData = variableListData;
    if (msItems.length) payload.manualSelectContentInput = { manualSelectContentItems: msItems };
    return payload;
  }

  return {
    scalars, setScalars,
    tables, setTables,
    vlScalars, setVlScalars,
    vlTables, setVlTables,
    grpInc, setGrpInc,
    extSel, setExtSel,
    buildPayload,
  };
}
