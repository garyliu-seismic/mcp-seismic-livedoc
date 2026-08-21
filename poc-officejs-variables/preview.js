'use strict';

// ── Preview — replace tokens and download PPTX ────────────────────────────────
//
// Flow: getFileAsync (sliced bytes) → JSZip → expand dynamic content →
// heal split-run tokens → regex-replace {{Var}} → re-zip → browser download.

async function previewDoc() {
  showStatus('Building preview…', 'info');

  try {
    // Step 1 — Scan all slides for dynamic configs (Office.js API calls)
    var dynamicTableMap = await scanDynamicTables();
    var dynamicImageMap = await scanDynamicImages();
    var dynamicChartMap = await scanDynamicCharts();

    // Step 2 — Build scalar/computed replacement map
    var replacements = {};
    variables.forEach(function (v) {
      if (v.kind === 'scalar') {
        replacements[v.name] = v.defaultValue || '';
      } else if (v.kind === 'computed') {
        replacements[v.name] = v.formula ? evaluateFormula(v.formula) : '';
      }
    });

    // Step 3 — Build table-variable payload for C# OOXML engine
    var tableVariables = (variables || [])
      .filter(function (v) { return v.kind === 'table' || v.kind === 'system'; })
      .map(function (v) {
        return {
          name: v.name,
          columns: (v.columns || []).map(function (c) { return String(c); }),
          rows: (v.rows || []).map(function (row) {
            return (row || []).map(function (cell) { return cell == null ? '' : String(cell); });
          }),
        };
      });

    // Step 4 — Send PPTX to C# preview engine
    var bytes = await getPptxBytes();
    var resp = await fetch('/api/preview-csharp', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        pptxBase64: bytesToBase64(bytes),
        payload: {
          replacements: replacements,
          dynamicTables: normalizeDynamicTablePayload(dynamicTableMap),
          dynamicImages: normalizeDynamicImagePayload(dynamicImageMap),
          dynamicCharts: normalizeDynamicChartPayload(dynamicChartMap),
          tableVariables: tableVariables,
        },
      }),
    });

    var body = await resp.json();
    if (!resp.ok) {
      throw new Error((body && (body.detail || body.error)) || 'C# preview failed');
    }

    var outBytes = base64ToBytes(body.pptxBase64 || '');
    downloadPptx(outBytes, 'preview.pptx');

    var sum = body.summary || {};
    var summary = 'Preview downloaded (C# OOXML)';
    if (typeof sum.expandedTables === 'number' || typeof sum.replacedTokens === 'number') {
      summary += ' — ' + (sum.expandedTables || 0) + ' dynamic table(s) expanded, ' +
                 (sum.replacedTokens || 0) + ' token(s) replaced';
      if (sum.slideCount) summary += ' across ' + sum.slideCount + ' slide(s)';
      summary += '.';
    } else {
      summary += '.';
    }
    showStatus(summary, 'success');

  } catch (err) {
    showStatus('Preview failed: ' + err.message, 'error');
    console.error('[Preview]', err);
  }
}

function normalizeDynamicTablePayload(map) {
  var out = {};
  Object.keys(map || {}).forEach(function (key) {
    var items = Array.isArray(map[key]) ? map[key] : [];
    out[String(key)] = items.map(function (cfg) {
      return {
        shapeName: cfg.shapeName,
        variableName: cfg.variableName,
        fromRow: cfg.fromRow,
        toRow: cfg.toRow,
        autoTokenize: cfg.autoTokenize || false,
        sortRules: cfg.sortRules || [],
        conditionalRows: cfg.conditionalRows || [],
        mergeRules: cfg.mergeRules || [],
        hideWhenEmpty: cfg.hideWhenEmpty || false,
      };
    });
  });
  return out;
}

function normalizeDynamicImagePayload(map) {
  var out = {};
  Object.keys(map || {}).forEach(function (key) {
    var items = Array.isArray(map[key]) ? map[key] : [];
    out[String(key)] = items.map(function (cfg) {
      return {
        shapeName: cfg.shapeName,
        fitMode: cfg.fitMode,
        sourceUrl: cfg.sourceUrl,
      };
    });
  });
  return out;
}

function normalizeDynamicChartPayload(map) {
  var out = {};
  Object.keys(map || {}).forEach(function (key) {
    var items = Array.isArray(map[key]) ? map[key] : [];
    out[String(key)] = items.map(function (cfg) {
      return {
        shapeName: cfg.shapeName,
        titleVar: cfg.titleVar,
        tableVar: cfg.tableVar,
        labelCol: cfg.labelCol,
        valueCol: cfg.valueCol,
      };
    });
  });
  return out;
}

function bytesToBase64(bytes) {
  var binary = '';
  var chunk = 0x8000;
  for (var i = 0; i < bytes.length; i += chunk) {
    var sub = bytes.subarray(i, i + chunk);
    binary += String.fromCharCode.apply(null, sub);
  }
  return btoa(binary);
}

function base64ToBytes(base64) {
  var binary = atob(base64);
  var out = new Uint8Array(binary.length);
  for (var i = 0; i < binary.length; i++) {
    out[i] = binary.charCodeAt(i);
  }
  return out;
}

// ── Heal split-run tokens ─────────────────────────────────────────────────────
//
// PowerPoint can split {{VarName}} like:
//   <a:r><a:t>{{Var</a:t></a:r><a:r><a:t>Name}}</a:t></a:r>
//
// Strategy: for each <a:p> paragraph, if the concatenated <a:t> text contains
// a token that is split across run boundaries, merge the minimal adjacent
// runs needed to make the token contiguous, preserving the first run's rPr.

function healSplitTokenRuns(xml, replacements) {
  if (!xml.includes('{{')) return xml;

  var parser = new DOMParser();
  var doc;
  try {
    doc = parser.parseFromString(xml, 'application/xml');
  } catch (e) {
    return xml;
  }
  if (doc.querySelector('parsererror')) return xml;

  var changed = false;
  var paragraphs = doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p');

  for (var pi = 0; pi < paragraphs.length; pi++) {
    var para    = paragraphs[pi];
    var runs    = Array.from(para.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'r'));
    if (runs.length < 2) continue;

    var texts = runs.map(function (r) {
      var tEl = r.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't')[0];
      return tEl ? tEl.textContent : '';
    });
    var combined = texts.join('');

    var tokenRe = /\{\{(\w+)\}\}/g;
    var needsMerge = false;
    var m;
    while ((m = tokenRe.exec(combined)) !== null) {
      var varName = m[1];
      if (!Object.prototype.hasOwnProperty.call(replacements, varName)) continue;
      var foundInSingleRun = texts.some(function (t) { return t.includes(m[0]); });
      if (!foundInSingleRun) { needsMerge = true; break; }
    }
    if (!needsMerge) continue;

    var firstRun = runs[0];
    var firstT   = firstRun.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't')[0];
    if (firstT) {
      firstT.textContent = combined;
      if (combined !== combined.trim()) {
        firstT.setAttribute('xml:space', 'preserve');
      }
    }
    for (var ri = 1; ri < runs.length; ri++) {
      runs[ri].parentNode.removeChild(runs[ri]);
    }
    changed = true;
  }

  if (!changed) return xml;

  var serializer = new XMLSerializer();
  var newXml = serializer.serializeToString(doc);
  newXml = newXml.replace(/ xmlns=""/g, '');
  return newXml;
}

// ── Get PPTX bytes from open document ─────────────────────────────────────────
function getPptxBytes() {
  return new Promise(function (resolve, reject) {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 262144 }, // 256 KB per slice
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error.message));
          return;
        }
        var file       = result.value;
        var sliceCount = file.sliceCount;
        var slices     = new Array(sliceCount);
        var remaining  = sliceCount;

        for (var i = 0; i < sliceCount; i++) {
          (function (idx) {
            file.getSliceAsync(idx, function (sliceResult) {
              if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                file.closeAsync();
                reject(new Error(sliceResult.error.message));
                return;
              }
              slices[sliceResult.value.index] = sliceResult.value.data;
              remaining--;
              if (remaining === 0) {
                file.closeAsync();
                var totalLen = slices.reduce(function (s, a) { return s + a.length; }, 0);
                var combined = new Uint8Array(totalLen);
                var offset   = 0;
                slices.forEach(function (s) { combined.set(s, offset); offset += s.length; });
                resolve(combined);
              }
            });
          })(i);
        }
      }
    );
  });
}

// ── Download helper ────────────────────────────────────────────────────────────
function downloadPptx(bytes, filename) {
  var blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  var url = URL.createObjectURL(blob);
  var a   = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
