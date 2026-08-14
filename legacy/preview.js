'use strict';

// ── Preview — replace tokens and download PPTX ────────────────────────────────
//
// Flow: getFileAsync (sliced bytes) → JSZip → expand dynamic content →
// heal split-run tokens → regex-replace {{Var}} → re-zip → browser download.

async function previewDoc() {
  if (typeof JSZip === 'undefined') {
    showStatus('JSZip library not loaded — check network connection.', 'error');
    return;
  }

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
      // table/system: handled separately — dynamic expansion or left as-is
    });

    // Step 3 — Get PPTX bytes
    var bytes = await getPptxBytes();
    var zip   = await JSZip.loadAsync(bytes);

    // Ordered slide paths: slide1.xml → index 1, slide2.xml → index 2, …
    var slidePaths = Object.keys(zip.files)
      .filter(function (n) { return /^ppt\/slides\/slide\d+\.xml$/.test(n); })
      .sort(function (a, b) {
        return parseInt(a.match(/slide(\d+)\.xml/)[1]) -
               parseInt(b.match(/slide(\d+)\.xml/)[1]);
      });

    if (!slidePaths.length) {
      showStatus('No slides found in the PPTX — try saving the document first.', 'error');
      return;
    }

    var replacedCount  = 0;
    var expandedTables = 0;
    var expandedImages = 0;
    var expandedCharts = 0;

    for (var si = 0; si < slidePaths.length; si++) {
      var slidePath    = slidePaths[si];
      var slideIndex   = si + 1; // 1-based
      var xml          = await zip.files[slidePath].async('string');

      // Pass A — expand dynamic table rows (before token replacement)
      var dynConfigs = dynamicTableMap[slideIndex] || [];
      if (dynConfigs.length) {
        var before = xml;
        xml = expandDynamicTablesInXml(xml, dynConfigs);
        if (xml !== before) expandedTables += dynConfigs.length;
      }

      // Pass A2 — replace dynamic image placeholders + strip indicator shapes
      var imgConfigs = dynamicImageMap[slideIndex] || [];
      xml = await expandDynamicImagesInXml(zip, xml, slideIndex, imgConfigs);
      if (imgConfigs.length) expandedImages += imgConfigs.length;

      // Pass A3 — update dynamic chart cached data (modifies chartN.xml files in zip)
      var chartConfigs = dynamicChartMap[slideIndex] || [];
      if (chartConfigs.length) {
        await expandDynamicChartsInXml(zip, xml, slideIndex, chartConfigs);
        expandedCharts += chartConfigs.length;
      }

      // Pass B — heal split-run tokens ({{Var}} fragmented across <a:r> runs)
      xml = healSplitTokenRuns(xml, replacements);

      // Pass C — regex replace remaining scalar/computed tokens
      xml = xml.replace(/\{\{(\w+)\}\}/g, function (match, varName) {
        if (Object.prototype.hasOwnProperty.call(replacements, varName)) {
          replacedCount++;
          return xmlEscapeVal(replacements[varName]);
        }
        return match;
      });

      zip.file(slidePath, xml);
    }

    // Re-zip and trigger download
    var newBytes = await zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });

    downloadPptx(newBytes, 'preview.pptx');

    var summary = 'Preview downloaded — ';
    if (expandedTables) summary += expandedTables + ' dynamic table(s) expanded, ';
    if (expandedImages) summary += expandedImages + ' dynamic image(s) replaced, ';
    if (expandedCharts) summary += expandedCharts + ' dynamic chart(s) updated, ';
    summary += replacedCount + ' token(s) replaced across ' + slidePaths.length + ' slide(s).';
    showStatus(summary, 'success');

  } catch (err) {
    showStatus('Preview failed: ' + err.message, 'error');
    console.error('[Preview]', err);
  }
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
