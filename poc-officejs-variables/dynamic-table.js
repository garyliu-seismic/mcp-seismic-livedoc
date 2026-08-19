'use strict';

// ── Dynamic Table Settings ─────────────────────────────────────────────────────
//
// A PPT table shape is "tagged" as dynamic by storing a JSON config in
// shape.tags under the key LIVEDOC_DYN_TABLE.  At preview time we:
//   1. Scan all slides via Office.js API to collect those configs.
//   2. In the JSZip pass, find the matching <p:graphicFrame> by shape name,
//      clone the template <a:tr> rows for every data record, then drop the
//      original template rows.

// Module-level state for the dynamic table drawer sections
var dynConditionalRows = [];
var dynMergeRules      = [];
var dynSortRules       = [];
var dynHideWhenEmpty   = false;

async function openDynamicDrawer() {
  editMode = 'dynamic';
  openDrawer('Dynamic Table Settings');
  document.getElementById('drawerForm').innerHTML =
    '<p class="form-hint" style="padding:8px 0">Loading table shapes from current slide…</p>';
  document.getElementById('drawerSaveBtn').textContent = 'Save Config';

  var shapes;
  try {
    shapes = await loadCurrentSlideTableShapes();
  } catch (e) {
    document.getElementById('drawerForm').innerHTML =
      '<p class="form-hint" style="color:#c00">Could not load slide shapes: ' + h(e.message) + '</p>';
    return;
  }
  buildDynamicDrawerForm(shapes);
}

async function loadCurrentSlideTableShapes() {
  var all = [], tables = [];
  await PowerPoint.run(async function (context) {
    var slide;
    try {
      var sel = context.presentation.getSelectedSlides();
      sel.load('items/id');
      await context.sync();
      slide = sel.items[0];
      currentSlideId = slide ? slide.id : null;
    } catch (_) {
      slide = context.presentation.slides.getItemAt(0);
    }
    slide.shapes.load('items/name,items/type');
    await context.sync();

    slide.shapes.items.forEach(function (s) {
      all.push({ name: s.name, type: s.type });
      if (s.type === 'Table' ||
          (typeof PowerPoint.ShapeType !== 'undefined' && s.type === PowerPoint.ShapeType.table)) {
        tables.push({ name: s.name });
      }
    });
  });
  // Fallback: if type-check found nothing, surface all shapes
  return tables.length ? tables : all.map(function (s) { return { name: s.name + ' (' + s.type + ')' }; });
}

function buildDynamicDrawerForm(tableShapes) {
  var form    = document.getElementById('drawerForm');
  var tableVars = variables.filter(function (v) { return v.kind === 'table'; });

  if (!tableShapes.length) {
    form.innerHTML = '<p class="form-hint">No table shapes found on the current slide.<br>Insert a PowerPoint table first, then open this dialog.</p>';
    return;
  }
  if (!tableVars.length) {
    form.innerHTML = '<p class="form-hint">No table variables defined yet.<br>Add a Table Variable first.</p>';
    return;
  }

  var shapeOpts = tableShapes.map(function (s) {
    var sel = (selectedTableName && s.name === selectedTableName) ? ' selected' : '';
    return '<option value="' + h(s.name) + '"' + sel + '>' + h(s.name) + '</option>';
  }).join('');

  var varOpts = tableVars.map(function (v) {
    return '<option value="' + h(v.name) + '">' + h(v.name) +
           ' (' + (v.columns || []).length + ' cols)</option>';
  }).join('');

  dynConditionalRows = [];
  dynMergeRules      = [];
  dynSortRules       = [];
  dynHideWhenEmpty   = false;

  form.innerHTML =
    '<div class="form-row"><label>Slide Table</label>' +
    '<select id="df-dyn-shape" class="field">' + shapeOpts + '</select></div>' +
    '<div class="form-row"><label>Variable</label>' +
    '<select id="df-dyn-var" class="field" onchange="renderConditionalRows();renderMergeRules();renderSortRules()">' + varOpts + '</select></div>' +
    '<div class="form-row"><label>Repeat Rows</label>' +
    '<div class="row-range-wrap">' +
    '<span class="range-label">From</span>' +
    '<input id="df-dyn-from" class="field field-sm" type="number" min="1" value="2" oninput="updateZoneLabel()"/>' +
    '<span class="range-label">To</span>' +
    '<input id="df-dyn-to"   class="field field-sm" type="number" min="1" value="2" oninput="updateZoneLabel()"/>' +
    '</div></div>' +
    '<div id="df-dyn-zone-hint" class="form-hint" style="font-style:italic;color:#777;margin-bottom:6px"></div>' +
    '<div class="form-section-hdr">Sort Rules <button class="btn-sm" onclick="addSortRule()">+ Add</button></div>' +
    '<div id="df-dyn-sort-list" style="margin-bottom:4px"></div>' +
    '<div class="form-section-hdr" style="margin-top:10px">Conditional Rows <button class="btn-sm" onclick="addConditionalRow()">+ Add</button></div>' +
    '<div id="df-dyn-cond-list" style="margin-bottom:4px"></div>' +
    '<div class="form-section-hdr" style="margin-top:10px">Merge Rules (Vertical) <button class="btn-sm" onclick="addMergeRule()">+ Add</button></div>' +
    '<div id="df-dyn-merge-list" style="margin-bottom:4px"></div>' +
    '<div style="margin-top:12px;display:flex;align-items:center;gap:6px">' +
    '<input type="checkbox" id="df-dyn-hide" onchange="dynHideWhenEmpty=this.checked"/>' +
    '<label for="df-dyn-hide" style="font-size:12px;color:#444;cursor:pointer">Hide slide when no data</label>' +
    '</div>';

  updateZoneLabel();
  renderSortRules();
  renderConditionalRows();
  renderMergeRules();
}

function getSelectedVariableColumns() {
  var varName = (document.getElementById('df-dyn-var') || {}).value || '';
  var tv = variables.find(function (v) { return v.name === varName && v.kind === 'table'; });
  return tv ? (tv.columns || []) : [];
}

function updateZoneLabel() {
  var fromRow = parseInt((document.getElementById('df-dyn-from') || {}).value || '2', 10);
  var toRow   = parseInt((document.getElementById('df-dyn-to')   || {}).value || '2', 10);
  var el = document.getElementById('df-dyn-zone-hint');
  if (!el) return;
  var parts = [];
  if (fromRow > 1) parts.push('Rows 1–' + (fromRow - 1) + ' = header');
  parts.push('Rows ' + fromRow + '–' + toRow + ' = repeating template');
  parts.push('Rows ' + (toRow + 1) + '+ = footer');
  el.textContent = parts.join('  ·  ');
}

function renderSortRules() {
  var el = document.getElementById('df-dyn-sort-list');
  if (!el) return;
  var cols = getSelectedVariableColumns();
  if (!dynSortRules.length) {
    el.innerHTML = '<p class="form-hint" style="margin:2px 0">No sort rules. Click + Add to sort rows before expansion.</p>';
    return;
  }
  el.innerHTML = dynSortRules.map(function (rule, i) {
    var colOpts = cols.map(function (c) {
      return '<option value="' + h(c) + '"' + (c === rule.column ? ' selected' : '') + '>' + h(c) + '</option>';
    }).join('');
    return '<div style="display:flex;align-items:center;gap:4px;margin-bottom:3px">' +
      '<select class="field" style="width:140px" onchange="dynSortRules[' + i + '].column = this.value">' + colOpts + '</select>' +
      '<select class="field" style="width:80px" onchange="dynSortRules[' + i + '].direction = this.value">' +
        '<option value="asc"'  + (rule.direction !== 'desc' ? ' selected' : '') + '>A → Z</option>' +
        '<option value="desc"' + (rule.direction === 'desc'  ? ' selected' : '') + '>Z → A</option>' +
      '</select>' +
      '<button class="btn-icon" onclick="removeSortRule(' + i + ')" title="Remove">✕</button>' +
    '</div>';
  }).join('');
}

function addSortRule() {
  var cols = getSelectedVariableColumns();
  dynSortRules.push({ column: cols[0] || '', direction: 'asc' });
  renderSortRules();
}

function removeSortRule(idx) {
  dynSortRules.splice(idx, 1);
  renderSortRules();
}

function renderConditionalRows() {
  var el = document.getElementById('df-dyn-cond-list');
  if (!el) return;
  var cols = getSelectedVariableColumns();
  if (!dynConditionalRows.length) {
    el.innerHTML = '<p class="form-hint" style="margin:2px 0">No conditional rows. Click + Add to include a row only when a column meets a condition.</p>';
    return;
  }
  el.innerHTML = dynConditionalRows.map(function (rule, i) {
    var colSel = cols.map(function (c) {
      return '<option value="' + h(c) + '"' + (c === rule.column ? ' selected' : '') + '>' + h(c) + '</option>';
    }).join('');
    var showVal = (rule.operator === 'equals' || rule.operator === 'notEquals');
    return '<div style="display:flex;align-items:center;gap:4px;margin-bottom:3px">' +
      '<span style="font-size:11px;color:#555">Row</span>' +
      '<input type="number" min="1" class="field field-sm" style="width:46px" value="' + (rule.rowIndex || '') + '" ' +
        'onchange="dynConditionalRows[' + i + '].rowIndex = parseInt(this.value,10)" placeholder="#"/>' +
      '<select class="field" style="width:100px" onchange="dynConditionalRows[' + i + '].column = this.value">' + colSel + '</select>' +
      '<select class="field" style="width:106px" onchange="onCondOpChange(' + i + ',this.value)">' +
        '<option value="notEmpty"' + (rule.operator === 'notEmpty' ? ' selected' : '') + '>is not empty</option>' +
        '<option value="empty"'   + (rule.operator === 'empty'    ? ' selected' : '') + '>is empty</option>' +
        '<option value="equals"'  + (rule.operator === 'equals'   ? ' selected' : '') + '>equals</option>' +
        '<option value="notEquals"'+(rule.operator === 'notEquals'? ' selected' : '') + '>not equals</option>' +
      '</select>' +
      '<input type="text" class="field field-sm" style="width:70px;' + (showVal ? '' : 'display:none;') + '" ' +
        'id="df-cond-val-' + i + '" value="' + h(rule.value || '') + '" ' +
        'onchange="dynConditionalRows[' + i + '].value = this.value" placeholder="Value"/>' +
      '<button class="btn-icon" onclick="removeConditionalRow(' + i + ')" title="Remove">✕</button>' +
    '</div>';
  }).join('');
}

function onCondOpChange(idx, op) {
  dynConditionalRows[idx].operator = op;
  var valEl = document.getElementById('df-cond-val-' + idx);
  if (valEl) valEl.style.display = (op === 'equals' || op === 'notEquals') ? '' : 'none';
}

function addConditionalRow() {
  var cols = getSelectedVariableColumns();
  var fromRow = parseInt((document.getElementById('df-dyn-from') || {}).value || '2', 10);
  dynConditionalRows.push({ rowIndex: fromRow, column: cols[0] || '', operator: 'notEmpty', value: '' });
  renderConditionalRows();
}

function removeConditionalRow(idx) {
  dynConditionalRows.splice(idx, 1);
  renderConditionalRows();
}

function renderMergeRules() {
  var el = document.getElementById('df-dyn-merge-list');
  if (!el) return;
  var cols = getSelectedVariableColumns();
  if (!dynMergeRules.length) {
    el.innerHTML = '<p class="form-hint" style="margin:2px 0">No merge rules. Click + Add to merge adjacent cells with the same value.</p>';
    return;
  }
  el.innerHTML = dynMergeRules.map(function (rule, i) {
    var colOpts = cols.map(function (c, ci) {
      var idx = ci + 1;
      return '<option value="' + idx + '"' + (idx === rule.columnIndex ? ' selected' : '') + '>' + h(c) + '</option>';
    }).join('');
    return '<div style="display:flex;align-items:center;gap:4px;margin-bottom:3px">' +
      '<span style="font-size:11px;color:#555">Column</span>' +
      '<select class="field" style="width:130px" onchange="dynMergeRules[' + i + '].columnIndex = parseInt(this.value,10)">' + colOpts + '</select>' +
      '<select class="field" style="width:110px" onchange="dynMergeRules[' + i + '].strategy = this.value">' +
        '<option value="sameValue"' + (rule.strategy === 'sameValue' ? ' selected' : '') + '>Same value</option>' +
        '<option value="empty"'     + (rule.strategy === 'empty'     ? ' selected' : '') + '>Empty cells</option>' +
      '</select>' +
      '<button class="btn-icon" onclick="removeMergeRule(' + i + ')" title="Remove">✕</button>' +
    '</div>';
  }).join('');
}

function addMergeRule() {
  dynMergeRules.push({ columnIndex: 1, strategy: 'sameValue' });
  renderMergeRules();
}

function removeMergeRule(idx) {
  dynMergeRules.splice(idx, 1);
  renderMergeRules();
}

// Reload the Dynamic table drawer for the current slide (called on slide change).
async function refreshDynamicDrawer() {
  var form = document.getElementById('drawerForm');
  if (!form) return;
  form.innerHTML = '<p class="form-hint" style="padding:8px 0">Loading table shapes from current slide…</p>';
  try {
    var shapes = await loadCurrentSlideTableShapes();
    buildDynamicDrawerForm(shapes);
  } catch (e) {
    form.innerHTML = '<p class="form-hint" style="color:#c00">Could not load slide shapes: ' + h(e.message) + '</p>';
  }
}

// Sync just the Slide Table dropdown to a given shape name.
function syncDynamicDrawerDropdown(tableName) {
  var sel = document.getElementById('df-dyn-shape');
  if (!sel || !tableName) return;
  for (var i = 0; i < sel.options.length; i++) {
    if (sel.options[i].value === tableName) { sel.selectedIndex = i; break; }
  }
}

// Sync the Variable dropdown, Repeat Rows fields, and advanced sections to an existing config (or reset to defaults).
function syncDynamicDrawerFromConfig(config) {
  var varSel = document.getElementById('df-dyn-var');
  var fromEl = document.getElementById('df-dyn-from');
  var toEl   = document.getElementById('df-dyn-to');
  if (!varSel || !fromEl || !toEl) return;

  if (config) {
    for (var i = 0; i < varSel.options.length; i++) {
      if (varSel.options[i].value === config.variableName) { varSel.selectedIndex = i; break; }
    }
    fromEl.value = config.fromRow || 1;
    toEl.value   = config.toRow   || 1;
    dynConditionalRows = (config.conditionalRows || []).map(function (r) { return Object.assign({}, r); });
    dynMergeRules      = (config.mergeRules      || []).map(function (r) { return Object.assign({}, r); });
    dynSortRules       = (config.sortRules        || []).map(function (r) { return Object.assign({}, r); });
    dynHideWhenEmpty   = !!config.hideWhenEmpty;
    var hideEl = document.getElementById('df-dyn-hide');
    if (hideEl) hideEl.checked = dynHideWhenEmpty;
  } else {
    varSel.selectedIndex = 0;
    fromEl.value = 2;
    toEl.value   = 2;
    dynConditionalRows = [];
    dynMergeRules      = [];
    dynSortRules       = [];
    dynHideWhenEmpty   = false;
    var hideEl = document.getElementById('df-dyn-hide');
    if (hideEl) hideEl.checked = false;
  }
  updateZoneLabel();
  renderSortRules();
  renderConditionalRows();
  renderMergeRules();
}

function saveDynamicTableConfig() {
  var shapeName = (document.getElementById('df-dyn-shape') || {}).value || '';
  var varName   = (document.getElementById('df-dyn-var')   || {}).value || '';
  var fromRow   = parseInt((document.getElementById('df-dyn-from') || {}).value || '1', 10);
  var toRow     = parseInt((document.getElementById('df-dyn-to')   || {}).value || '1', 10);

  if (!shapeName) { showStatus('Select a slide table shape.', 'error'); return; }
  if (!varName)   { showStatus('Select a table variable.', 'error');    return; }
  if (isNaN(fromRow) || isNaN(toRow) || fromRow < 1 || toRow < fromRow) {
    showStatus('Repeating row range is invalid (From must be ≤ To).', 'error'); return;
  }

  var config = {
    variableName: varName,
    fromRow: fromRow,
    toRow: toRow,
    sortRules: dynSortRules.slice(),
    conditionalRows: dynConditionalRows.slice(),
    mergeRules: dynMergeRules.slice(),
    hideWhenEmpty: dynHideWhenEmpty,
  };

  PowerPoint.run(async function (context) {
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
    if (!shape) {
      showStatus('Shape "' + shapeName + '" not found on current slide.', 'error');
      return;
    }

    shape.tags.add('LIVEDOC_DYN_TABLE', JSON.stringify(config));
    await context.sync();

    closeDrawer();
    showStatus(
      '"' + shapeName + '" → variable "' + varName +
      '", template rows ' + fromRow + '–' + toRow + '. Run Preview to expand.',
      'success'
    );
  }).catch(function (e) {
    showStatus('Failed to tag shape: ' + e.message, 'error');
  });
}

// ── Dynamic Table — scan & expand ─────────────────────────────────────────────

// Pre-scan all slides for dynamic table shape tags (called at preview start).
// Returns: { slideIndex(1-based): [{shapeName, variableName, fromRow, toRow}] }
async function scanDynamicTables() {
  var result = {};

  await PowerPoint.run(async function (context) {
    var slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    slides.items.forEach(function (slide) {
      slide.shapes.load('items/name,items/type');
    });
    await context.sync();

    var shapesBySlide = [];
    slides.items.forEach(function (slide, si) {
      var allShapes = slide.shapes.items || [];
      if (allShapes.length) shapesBySlide.push({ slideIndex: si + 1, shapes: allShapes });
    });

    if (!shapesBySlide.length) return;

    shapesBySlide.forEach(function (entry) {
      entry.shapes.forEach(function (s) { s.tags.load('items/key,items/value'); });
    });
    await context.sync();

    shapesBySlide.forEach(function (entry) {
      entry.shapes.forEach(function (s) {
        var dynTag = s.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_TABLE'; });
        if (!dynTag) return;
        try {
          var config = JSON.parse(dynTag.value);
          if (!result[entry.slideIndex]) result[entry.slideIndex] = [];
          result[entry.slideIndex].push(Object.assign({ shapeName: s.name }, config));
        } catch (_) {}
      });
    });
  });

  return result;
}

// Expand dynamic table rows inside a single slide's XML string.
// configs: [{shapeName, variableName, fromRow, toRow}]  (fromRow/toRow are 1-indexed)
function expandDynamicTablesInXml(xml, configs) {
  if (!configs || !configs.length) return xml;

  var NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  var parser = new DOMParser();
  var doc    = parser.parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parseerror').length) return xml;

  var changed = false;

  configs.forEach(function (config) {
    var tv = variables.find(function (v) {
      return v.name === config.variableName && v.kind === 'table';
    });
    if (!tv || !(tv.rows || []).length) return;

    // Locate the <p:graphicFrame> whose <p:cNvPr name="shapeName"> matches
    var nvFramePrs = doc.getElementsByTagNameNS(NS_P, 'nvGraphicFramePr');
    var targetTbl  = null;

    for (var fi = 0; fi < nvFramePrs.length; fi++) {
      var nvPr   = nvFramePrs[fi];
      var cNvPrs = nvPr.childNodes;
      var found  = false;
      for (var ci = 0; ci < cNvPrs.length; ci++) {
        var child = cNvPrs[ci];
        if (child.nodeType === 1 && child.localName === 'cNvPr' &&
            child.getAttribute('name') === config.shapeName) {
          found = true; break;
        }
      }
      if (!found) continue;

      // Walk up to the <p:graphicFrame>, then find <a:tbl>
      var frame = nvPr.parentNode;
      var tbls  = frame.getElementsByTagNameNS(NS_A, 'tbl');
      if (tbls.length) { targetTbl = tbls[0]; break; }
    }

    if (!targetTbl) {
      console.warn('[DynTable] Shape "' + config.shapeName + '" not found in slide XML.');
      return;
    }

    // Direct <a:tr> children only (not nested)
    var allRows = Array.from(targetTbl.childNodes).filter(function (n) {
      return n.nodeType === 1 && n.localName === 'tr';
    });
    if (!allRows.length) return;

    var fromIdx = (config.fromRow || 1) - 1; // convert to 0-indexed
    var toIdx   = (config.toRow   || 1) - 1;
    if (fromIdx < 0 || toIdx >= allRows.length || fromIdx > toIdx) {
      console.warn('[DynTable] Row range ' + config.fromRow + '-' + config.toRow +
                   ' out of bounds (table has ' + allRows.length + ' rows).');
      return;
    }

    var templateRows = allRows.slice(fromIdx, toIdx + 1);
    var insertBefore = templateRows[0];

    // For each data record: clone each template row and substitute {{ColName}}
    tv.rows.forEach(function (dataRow) {
      var rowMap = {};
      (tv.columns || []).forEach(function (col, i) { rowMap[col] = (dataRow || [])[i] || ''; });

      templateRows.forEach(function (templateRow) {
        var newRow = templateRow.cloneNode(true);

        var tEls = newRow.getElementsByTagNameNS(NS_A, 't');
        for (var ti = 0; ti < tEls.length; ti++) {
          var tEl = tEls[ti];
          tEl.textContent = tEl.textContent.replace(/\{\{(\w+)\}\}/g, function (match, colName) {
            return Object.prototype.hasOwnProperty.call(rowMap, colName) ? rowMap[colName] : match;
          });
        }
        targetTbl.insertBefore(newRow, insertBefore);
      });
    });

    // Remove the original template rows
    templateRows.forEach(function (r) { targetTbl.removeChild(r); });
    changed = true;
  });

  if (!changed) return xml;

  var serializer = new XMLSerializer();
  var result = serializer.serializeToString(doc);
  result = result.replace(/ xmlns=""/g, '');
  return result;
}
