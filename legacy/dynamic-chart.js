'use strict';

// ── Dynamic Chart Settings ─────────────────────────────────────────────────────
//
// A slide chart shape is tagged with LIVEDOC_DYN_CHART: { titleVar, tableVar, labelCol, valueCol }
// The chart shape itself stays on the slide (the sample data is the visual placeholder).
// At preview time we navigate: slide XML → slide rels → ppt/charts/chartN.xml
// and replace cached chart data (title, <c:strCache> category labels, <c:numCache> values).
//
// UI mirrors the COM add-in: Binding Table Variable + Chart Title Variable +
// a series grid row "(Category)" with Legend (label col) and Data (value col) pickers.

async function openDynamicChartDrawer() {
  editMode = 'dynchart';
  openDrawer('Dynamic Chart Settings');
  document.getElementById('drawerForm').innerHTML =
    '<p class="form-hint" style="padding:8px 0">Loading shapes from current slide…</p>';
  document.getElementById('drawerSaveBtn').textContent = 'Save Config';

  var shapes;
  try {
    shapes = await loadCurrentSlideAllShapes();
  } catch (e) {
    document.getElementById('drawerForm').innerHTML =
      '<p class="form-hint" style="color:#c00">Could not load shapes: ' + h(e.message) + '</p>';
    return;
  }
  buildDynamicChartDrawerForm(shapes);
}

async function refreshDynamicChartDrawer() {
  var form = document.getElementById('drawerForm');
  if (!form) return;
  form.innerHTML = '<p class="form-hint" style="padding:8px 0">Loading shapes from current slide…</p>';
  try {
    var shapes = await loadCurrentSlideAllShapes();
    buildDynamicChartDrawerForm(shapes);
  } catch (e) {
    form.innerHTML = '<p class="form-hint" style="color:#c00">Could not load shapes: ' + h(e.message) + '</p>';
  }
}

function syncDynamicChartDrawerDropdown(shapeName) {
  var sel = document.getElementById('df-chart-shape');
  if (!sel || !shapeName) return;
  for (var i = 0; i < sel.options.length; i++) {
    if (sel.options[i].value === shapeName) { sel.selectedIndex = i; break; }
  }
}

// Restore all chart drawer dropdowns from a previously saved config.
function syncDynamicChartDrawerFromConfig(config) {
  if (!config) return;

  var titleSel = document.getElementById('df-chart-title');
  if (titleSel && config.titleVar) {
    for (var i = 0; i < titleSel.options.length; i++) {
      if (titleSel.options[i].value === config.titleVar) { titleSel.selectedIndex = i; break; }
    }
  }

  var tableSel = document.getElementById('df-chart-table');
  if (tableSel && config.tableVar) {
    for (var j = 0; j < tableSel.options.length; j++) {
      if (tableSel.options[j].value === config.tableVar) { tableSel.selectedIndex = j; break; }
    }
    onChartTableChanged();
  }

  var labelSel = document.getElementById('df-chart-label');
  if (labelSel && config.labelCol) {
    for (var k = 0; k < labelSel.options.length; k++) {
      if (labelSel.options[k].value === config.labelCol) { labelSel.selectedIndex = k; break; }
    }
  }

  var valueSel = document.getElementById('df-chart-value');
  if (valueSel && config.valueCol) {
    for (var l = 0; l < valueSel.options.length; l++) {
      if (valueSel.options[l].value === config.valueCol) { valueSel.selectedIndex = l; break; }
    }
  }
}

function buildDynamicChartDrawerForm(shapes) {
  var form       = document.getElementById('drawerForm');
  var tableVars  = variables.filter(function (v) { return v.kind === 'table'; });
  var scalarVars = variables.filter(function (v) { return v.kind === 'scalar'; });

  if (!shapes.length) {
    form.innerHTML = '<p class="form-hint">No shapes found on the current slide.</p>';
    return;
  }
  if (!tableVars.length) {
    form.innerHTML = '<p class="form-hint">No table variables defined. Add a Table Variable first.</p>';
    return;
  }

  var shapeOpts = shapes.map(function (s) {
    var sel = (selectedChartName && s.name === selectedChartName) ? ' selected' : '';
    return '<option value="' + h(s.name) + '"' + sel + '>' + h(s.name) + ' (' + h(s.type) + ')</option>';
  }).join('');

  var titleOpts = '<option value="">(none)</option>' + scalarVars.map(function (v) {
    return '<option value="' + h(v.name) + '">' + h(v.name) + '</option>';
  }).join('');

  var tableOpts = tableVars.map(function (v) {
    return '<option value="' + h(v.name) + '">' + h(v.name) +
           ' (' + (v.columns || []).length + ' cols)</option>';
  }).join('');

  var firstTable = tableVars[0];
  var colOpts = (firstTable.columns || []).map(function (c) {
    return '<option value="' + h(c) + '">' + h(c) + '</option>';
  }).join('');
  var valueColOpts = (firstTable.columns || []).map(function (c, i) {
    var sel = (i === 1) ? ' selected' : '';
    return '<option value="' + h(c) + '"' + sel + '>' + h(c) + '</option>';
  }).join('');

  form.innerHTML =
    '<div class="form-row"><label>Chart Shape</label>' +
    '<select id="df-chart-shape" class="field">' + shapeOpts + '</select></div>' +
    '<div class="form-row"><label>Binding Table Variable</label>' +
    '<select id="df-chart-table" class="field" onchange="onChartTableChanged()">' + tableOpts + '</select></div>' +
    '<div class="form-row"><label>Chart Title Variable</label>' +
    '<select id="df-chart-title" class="field">' + titleOpts + '</select></div>' +
    '<div class="chart-series-grid">' +
    '<div class="chart-series-header">' +
    '<span class="csg-name">Name</span>' +
    '<span class="csg-legend">Legend <span class="opt">(label col)</span></span>' +
    '<span class="csg-data">Data <span class="opt">(value col)</span></span>' +
    '</div>' +
    '<div class="chart-series-row">' +
    '<span class="csg-name">(Category)</span>' +
    '<select id="df-chart-label" class="field csg-legend">' + colOpts + '</select>' +
    '<select id="df-chart-value" class="field csg-data">' + valueColOpts + '</select>' +
    '</div></div>' +
    '<p class="form-hint">Click the chart shape on the slide to auto-select it above. ' +
    'At preview, title and pie data are replaced with variable values.</p>';
}

// Called when the Binding Table Variable dropdown changes — refreshes column pickers.
function onChartTableChanged() {
  var tableName = (document.getElementById('df-chart-table') || {}).value || '';
  var tv = variables.find(function (v) { return v.name === tableName && v.kind === 'table'; });
  var cols = tv ? (tv.columns || []) : [];

  var colOpts = cols.map(function (c) {
    return '<option value="' + h(c) + '">' + h(c) + '</option>';
  }).join('');
  var valueColOpts = cols.map(function (c, i) {
    var sel = (i === 1) ? ' selected' : '';
    return '<option value="' + h(c) + '"' + sel + '>' + h(c) + '</option>';
  }).join('');

  var labelSel = document.getElementById('df-chart-label');
  var valueSel = document.getElementById('df-chart-value');
  if (labelSel) labelSel.innerHTML = colOpts;
  if (valueSel) valueSel.innerHTML = valueColOpts;
}

async function saveDynamicChartConfig() {
  var shapeName = (document.getElementById('df-chart-shape') || {}).value || '';
  var tableVar  = (document.getElementById('df-chart-table') || {}).value || '';
  var titleVar  = (document.getElementById('df-chart-title') || {}).value || '';
  var labelCol  = (document.getElementById('df-chart-label') || {}).value || '';
  var valueCol  = (document.getElementById('df-chart-value') || {}).value || '';

  if (!shapeName) { showStatus('Select a chart shape.', 'error'); return; }
  if (!tableVar)  { showStatus('Select a binding table variable.', 'error'); return; }
  if (!labelCol)  { showStatus('Select a Legend (label) column.', 'error'); return; }
  if (!valueCol)  { showStatus('Select a Data (value) column.', 'error'); return; }

  var config = { titleVar: titleVar || '', tableVar: tableVar, labelCol: labelCol, valueCol: valueCol };

  try {
    await PowerPoint.run(async function (context) {
      var slide;
      try {
        var sel = context.presentation.getSelectedSlides();
        sel.load('items');
        await context.sync();
        slide = sel.items[0];
      } catch (_) {
        slide = context.presentation.slides.getItemAt(0);
      }

      slide.shapes.load('items/name');
      await context.sync();

      var shape = slide.shapes.items.find(function (s) { return s.name === shapeName; });
      if (!shape) { showStatus('Shape "' + shapeName + '" not found on current slide.', 'error'); return; }

      shape.tags.add('LIVEDOC_DYN_CHART', JSON.stringify(config));
      await context.sync();

      closeDrawer();
      showStatus(
        '"' + shapeName + '" linked to table "' + tableVar +
        '"' + (titleVar ? ', title → "' + titleVar + '"' : '') +
        '. Run Preview to update chart data.',
        'success'
      );
    });
  } catch (e) {
    showStatus('Failed to tag chart: ' + e.message, 'error');
  }
}

// ── Dynamic Chart — scan & expand ─────────────────────────────────────────────

async function scanDynamicCharts() {
  var result = {};

  await PowerPoint.run(async function (context) {
    var slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    slides.items.forEach(function (slide) {
      slide.shapes.load('items/name');
    });
    await context.sync();

    var shapesBySlide = [];
    slides.items.forEach(function (slide, si) {
      var nonInd = slide.shapes.items.filter(function (s) {
        return !s.name.startsWith('__LIVEDOC_IND_');
      });
      if (nonInd.length) shapesBySlide.push({ slideIndex: si + 1, shapes: nonInd });
    });

    if (!shapesBySlide.length) return;

    shapesBySlide.forEach(function (entry) {
      entry.shapes.forEach(function (s) { s.tags.load('items/key,items/value'); });
    });
    await context.sync();

    shapesBySlide.forEach(function (entry) {
      entry.shapes.forEach(function (s) {
        var tag = s.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_CHART'; });
        if (!tag) return;
        try {
          var config = JSON.parse(tag.value);
          if (!result[entry.slideIndex]) result[entry.slideIndex] = [];
          result[entry.slideIndex].push(Object.assign({ shapeName: s.name }, config));
        } catch (_) {}
      });
    });
  });

  return result;
}

// Modify ppt/charts/chartN.xml in-place for each tagged chart on this slide.
// Returns the slide XML unchanged (only chart files are modified in the zip).
async function expandDynamicChartsInXml(zip, xml, slideIndex, configs) {
  if (!configs || !configs.length) return xml;

  var NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  var NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  var slideDoc = new DOMParser().parseFromString(xml, 'application/xml');
  if (slideDoc.getElementsByTagName('parseerror').length) return xml;

  var relsPath = 'ppt/slides/_rels/slide' + slideIndex + '.xml.rels';
  var relsXml  = zip.files[relsPath] ? await zip.files[relsPath].async('string') : '';

  for (var ci = 0; ci < configs.length; ci++) {
    var cfg = configs[ci];

    var rId = findChartRIdInSlideDoc(slideDoc, cfg.shapeName, NS_P, NS_C, NS_R);
    if (!rId) {
      console.warn('[DynChart] Chart shape "' + cfg.shapeName + '" not found in slide XML.');
      continue;
    }

    var relTarget = findRelTargetById(relsXml, rId);
    if (!relTarget) {
      console.warn('[DynChart] Rel "' + rId + '" not found in slide rels.');
      continue;
    }
    var chartPath = resolveRelativePath('ppt/slides/', relTarget);

    if (!zip.files[chartPath]) {
      console.warn('[DynChart] Chart file "' + chartPath + '" not in zip.');
      continue;
    }

    var chartXml = await zip.files[chartPath].async('string');

    var titleValue  = null;
    var labelValues = null;
    var numValues   = null;

    if (cfg.titleVar) {
      var tv = variables.find(function (v) { return v.name === cfg.titleVar && v.kind === 'scalar'; });
      if (tv) titleValue = tv.defaultValue || '';
    }

    if (cfg.tableVar) {
      var tbl = variables.find(function (v) { return v.name === cfg.tableVar && v.kind === 'table'; });
      if (tbl) {
        var labelIdx = (tbl.columns || []).indexOf(cfg.labelCol);
        var valueIdx = (tbl.columns || []).indexOf(cfg.valueCol);
        var rows = tbl.rows || [];
        if (labelIdx >= 0) labelValues = rows.map(function (r) { return (r || [])[labelIdx] || ''; });
        if (valueIdx >= 0) numValues   = rows.map(function (r) { return (r || [])[valueIdx] || '0'; });
      }
    }

    var updatedChartXml = updateChartCachedData(chartXml, titleValue, labelValues, numValues);
    zip.file(chartPath, updatedChartXml);

    await updateChartEmbeddedExcel(zip, chartPath, chartXml, labelValues, numValues);
  }

  return xml;
}

// Find the r:id attribute on <c:chart> inside a <p:graphicFrame>.
// Falls back to the first chart in the slide if no name match is found.
function findChartRIdInSlideDoc(slideDoc, shapeName, NS_P, NS_C, NS_R) {
  var frames = slideDoc.getElementsByTagNameNS(NS_P, 'graphicFrame');
  var firstChartRId = null;

  for (var i = 0; i < frames.length; i++) {
    var frame    = frames[i];
    var chartEls = frame.getElementsByTagNameNS(NS_C, 'chart');
    if (!chartEls.length) continue;

    var rId = chartEls[0].getAttributeNS(NS_R, 'id') || chartEls[0].getAttribute('r:id');
    if (!firstChartRId && rId) firstChartRId = rId;

    var nvPr  = frame.getElementsByTagNameNS(NS_P, 'nvGraphicFramePr')[0];
    if (!nvPr) continue;
    var cNvPr = findChildByLocalName(nvPr, 'cNvPr');
    if (cNvPr && cNvPr.getAttribute('name') === shapeName) return rId || null;
  }

  if (firstChartRId) {
    console.warn('[DynChart] Shape "' + shapeName + '" not found as graphicFrame; ' +
                 'falling back to first chart in slide (rId=' + firstChartRId + ').');
  }
  return firstChartRId;
}

// Update the embedded Excel workbook so "Edit Data" in PowerPoint shows the same data
// as the chart visual.  The XLSX is a nested ZIP inside the PPTX.
async function updateChartEmbeddedExcel(zip, chartPath, chartXml, labelValues, numValues) {
  if (!labelValues || !numValues || !labelValues.length) return;

  var extMatch = chartXml.match(/externalData[^>]+r:id="([^"]+)"/);
  if (!extMatch) extMatch = chartXml.match(/externalData[^>]+id="([^"]+)"/);
  if (!extMatch) return;
  var extRId = extMatch[1];

  var chartFilename = chartPath.substring(chartPath.lastIndexOf('/') + 1);
  var chartDir      = chartPath.substring(0, chartPath.lastIndexOf('/') + 1);
  var chartRelsPath = chartDir + '_rels/' + chartFilename + '.rels';
  if (!zip.files[chartRelsPath]) return;

  var chartRelsXml = await zip.files[chartRelsPath].async('string');
  var excelTarget  = findRelTargetById(chartRelsXml, extRId);
  if (!excelTarget) return;

  var excelPath = resolveRelativePath(chartDir, excelTarget);
  if (!zip.files[excelPath]) return;

  var excelBytes = await zip.files[excelPath].async('uint8array');
  var excelZip;
  try { excelZip = await JSZip.loadAsync(excelBytes); } catch (e) { return; }

  var wsPath = 'xl/worksheets/sheet1.xml';
  if (!excelZip.files[wsPath]) return;

  var wsXml = await excelZip.files[wsPath].async('string');
  excelZip.file(wsPath, updateExcelWorksheetData(wsXml, labelValues, numValues));

  var tablePath = 'xl/tables/table1.xml';
  if (excelZip.files[tablePath]) {
    var tableXml = await excelZip.files[tablePath].async('string');
    var newEnd = 'A1:D' + (1 + labelValues.length);
    excelZip.file(tablePath, tableXml.replace(/\bref="[^"]+"/g, 'ref="' + newEnd + '"'));
  }

  var newExcelBytes = await excelZip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
  zip.file(excelPath, newExcelBytes);
}

// Rebuild rows 2+ in a SpreadsheetML worksheet with our variable data.
// Category labels use t="inlineStr" to avoid touching sharedStrings.xml.
function updateExcelWorksheetData(wsXml, labelValues, numValues) {
  var NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var doc = new DOMParser().parseFromString(wsXml, 'application/xml');
  if (doc.getElementsByTagName('parseerror').length) return wsXml;

  var dim = doc.getElementsByTagNameNS(NS, 'dimension')[0];
  if (dim) dim.setAttribute('ref', 'A1:D' + (1 + labelValues.length));

  var sheetData = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
  if (!sheetData) return wsXml;

  Array.from(sheetData.getElementsByTagNameNS(NS, 'row')).forEach(function (r) {
    if (parseInt(r.getAttribute('r') || '1', 10) >= 2) sheetData.removeChild(r);
  });

  var count = Math.min(labelValues.length, numValues.length);
  for (var i = 0; i < count; i++) {
    var rn  = i + 2;
    var lbl = xmlEscapeVal(String(labelValues[i] || ''));
    var val = parseFloat(numValues[i]) || 0;
    var rowXml =
      '<row xmlns="' + NS + '" r="' + rn + '">' +
        '<c r="A' + rn + '" t="inlineStr"><is><t>' + lbl + '</t></is></c>' +
        '<c r="B' + rn + '"><v>' + val + '</v></c>' +
      '</row>';
    var rowDoc = new DOMParser().parseFromString(rowXml, 'application/xml');
    if (!rowDoc.getElementsByTagName('parseerror').length) {
      sheetData.appendChild(doc.importNode(rowDoc.documentElement, true));
    }
  }

  return new XMLSerializer().serializeToString(doc);
}

// Update chart data using literal elements (<c:strLit>/<c:numLit>) instead of cached refs.
// <c:strRef>/<c:numRef> point to the embedded Excel workbook — PowerPoint loads from there,
// ignoring the <c:strCache>/<c:numCache> we previously tried to update.
function updateChartCachedData(chartXml, titleValue, labelValues, numValues) {
  var NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  var doc = new DOMParser().parseFromString(chartXml, 'application/xml');
  if (doc.getElementsByTagName('parseerror').length) return chartXml;

  var changed = false;

  // 1. Update chart title
  if (titleValue !== null && titleValue !== '') {
    var titleEls = doc.getElementsByTagNameNS(NS_C, 'title');
    if (titleEls.length) {
      // Case A: rich-text title — find existing <a:t> and update it
      var tEls = titleEls[0].getElementsByTagNameNS(NS_A, 't');
      if (tEls.length) {
        tEls[0].textContent = titleValue;
        for (var ti = 1; ti < tEls.length; ti++) tEls[ti].textContent = '';
        changed = true;
      } else {
        // Case B: linked title (<c:strRef>) or empty rich-text — no <a:t> exists.
        // Replace the entire <c:tx> content with a new <c:rich> element so the
        // title is written as inline text instead of an Excel cell reference.
        var txEls = titleEls[0].getElementsByTagNameNS(NS_C, 'tx');
        var richXml = '<c:rich xmlns:c="' + NS_C + '" xmlns:a="' + NS_A + '">' +
          '<a:bodyPr/><a:lstStyle/>' +
          '<a:p><a:r><a:rPr lang="en-US" dirty="0"/>' +
          '<a:t>' + xmlEscapeVal(titleValue) + '</a:t>' +
          '</a:r></a:p></c:rich>';
        var richEl = importXmlFragment(doc, richXml);
        if (richEl) {
          if (txEls.length) {
            while (txEls[0].firstChild) txEls[0].removeChild(txEls[0].firstChild);
            txEls[0].appendChild(richEl);
          } else {
            var txEl = doc.createElementNS(NS_C, 'c:tx');
            txEl.appendChild(richEl);
            titleEls[0].insertBefore(txEl, titleEls[0].firstChild);
          }
          changed = true;
        }
      }
    }
  }

  // 2. Find first <c:ser> (pie chart has exactly one series)
  var serEls = doc.getElementsByTagNameNS(NS_C, 'ser');
  if (!serEls.length) {
    return changed ? new XMLSerializer().serializeToString(doc).replace(/ xmlns=""/g, '') : chartXml;
  }
  var ser = serEls[0];

  // 2a. Categories: replace <c:strRef> with <c:strLit> (inline literal, no Excel reference)
  if (labelValues && labelValues.length) {
    var catEl = ser.getElementsByTagNameNS(NS_C, 'cat')[0];
    if (catEl) {
      var strLitXml = '<c:strLit xmlns:c="' + NS_C + '">' +
        '<c:ptCount val="' + labelValues.length + '"/>' +
        labelValues.map(function (v, i) {
          return '<c:pt idx="' + i + '"><c:v>' + xmlEscapeVal(String(v)) + '</c:v></c:pt>';
        }).join('') + '</c:strLit>';
      var strLitEl = importXmlFragment(doc, strLitXml);
      if (strLitEl) {
        while (catEl.firstChild) catEl.removeChild(catEl.firstChild);
        catEl.appendChild(strLitEl);
        changed = true;
      }
    }
  }

  // 2b. Values: replace <c:numRef> with <c:numLit> (inline literal)
  if (numValues && numValues.length) {
    var valEl = ser.getElementsByTagNameNS(NS_C, 'val')[0];
    if (valEl) {
      var numLitXml = '<c:numLit xmlns:c="' + NS_C + '">' +
        '<c:formatCode>General</c:formatCode>' +
        '<c:ptCount val="' + numValues.length + '"/>' +
        numValues.map(function (v, i) {
          return '<c:pt idx="' + i + '"><c:v>' + String(parseFloat(v) || 0) + '</c:v></c:pt>';
        }).join('') + '</c:numLit>';
      var numLitEl = importXmlFragment(doc, numLitXml);
      if (numLitEl) {
        while (valEl.firstChild) valEl.removeChild(valEl.firstChild);
        valEl.appendChild(numLitEl);
        changed = true;
      }
    }
  }

  if (!changed) return chartXml;
  return new XMLSerializer().serializeToString(doc).replace(/ xmlns=""/g, '');
}
