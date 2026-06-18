'use strict';

// ── State ──────────────────────────────────────────────────────────────────────
/** @type {Array<{id:string, kind:'scalar'|'table'|'computed'|'system', name:string, group:string, [key:string]:any}>} */
let variables = [];

const expandedIds = new Set(); // IDs of table/system rows currently expanded
let searchQuery   = '';
let editMode      = 'add';    // 'add' | 'edit'
let editKind      = 'scalar';
let editId        = null;

// Tracks the name of the currently selected table shape in the slide (updated on selection change)
let selectedTableName   = null;
let selectedShapeName   = null;   // any non-indicator shape — used by dynimage drawer
let currentSlideId      = null;   // Tracks current slide; used to detect slide switches
let pendingImageDataUrl = null;   // file uploaded in the dynimage drawer; survives form rebuilds

const XML_NS    = 'http://schemas.livedoc.seismic.com/poc-variables/v2';
const XML_NS_V1 = 'http://schemas.livedoc.seismic.com/poc-variables/v1';

const TYPE_LABELS = { STRING: 'ABC', NUMBER: '123', DATE: 'DT', BOOLEAN: 'T/F' };

// Predefined system variable blueprints
const SYSTEM_PRESETS = {
  TOCEntries: {
    columns:     ['EntryName', 'EntryLevel', 'PageNumber', 'Index'],
    columnTypes: ['STRING', 'NUMBER', 'NUMBER', 'NUMBER'],
    description: 'Table of contents entries'
  }
};

// ── Office init ────────────────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('app').style.display     = 'flex';
    initSearch();
    initAddMenu();
    loadFromCustomXml();
    initSelectionHandler();
  } else {
    document.getElementById('loading').innerHTML =
      '<p style="color:#c00;padding:20px">This add-in requires PowerPoint.</p>';
  }
});

// ── Search ─────────────────────────────────────────────────────────────────────
function initSearch() {
  document.getElementById('searchInput').addEventListener('input', function () {
    searchQuery = this.value.trim().toLowerCase();
    renderTree();
  });
}

// ── Add-menu dropdown ──────────────────────────────────────────────────────────
function initAddMenu() {
  const btn  = document.getElementById('addBtn');
  const menu = document.getElementById('addMenu');
  btn.addEventListener('click', function (e) {
    e.stopPropagation();
    menu.classList.toggle('hidden');
  });
  document.addEventListener('click', function () {
    menu.classList.add('hidden');
  });
}

function openAddDrawer(kind) {
  document.getElementById('addMenu').classList.add('hidden');
  editMode = 'add';
  editKind = kind;
  editId   = null;
  buildDrawerForm(kind, null);
  openDrawer('Add ' + kindLabel(kind));
}

// ── Custom XML — Load ──────────────────────────────────────────────────────────
async function loadFromCustomXml() {
  showStatus('Loading from document…', 'info');
  try {
    await PowerPoint.run(async (ctx) => {
      const parts = ctx.presentation.customXmlParts;
      parts.load('items');
      await ctx.sync();

      if (!parts.items.length) {
        renderTree();
        showStatus('No saved variables in this document yet.', 'info');
        return;
      }

      const xmlResults = parts.items.map(p => p.getXml());
      await ctx.sync();
      const xmlStrings = xmlResults.map(r => r.value || '');

      let matchXml = xmlStrings.find(s => s.includes(XML_NS));
      let isV1     = false;
      if (!matchXml) {
        matchXml = xmlStrings.find(s => s.includes(XML_NS_V1));
        isV1     = !!matchXml;
      }

      if (!matchXml) {
        renderTree();
        showStatus('No saved variables in this document yet.', 'info');
        return;
      }

      try {
        if (isV1) { parseXmlV1(matchXml); }
        else       { parseXmlV2(matchXml); }
        showStatus('Loaded ' + variables.length + ' variable(s).', 'success');
      } catch (e) {
        showStatus('Parse error: ' + e.message, 'error');
      }
      renderTree();
    });
  } catch (e) {
    showStatus('Load error: ' + e.message, 'error');
    renderTree();
  }
}

function refreshFromDoc() {
  variables = [];
  expandedIds.clear();
  loadFromCustomXml();
}

// ── Custom XML — Save ──────────────────────────────────────────────────────────
async function saveToCustomXml(msg) {
  const xml = buildXmlV2();
  try {
    await PowerPoint.run(async (ctx) => {
      const parts = ctx.presentation.customXmlParts;
      parts.load('items');
      await ctx.sync();

      const xmlResults = parts.items.map(p => p.getXml());
      await ctx.sync();

      // Collect indices of old v1/v2 parts to remove
      const toDelete = [];
      xmlResults.forEach(function (r, i) {
        const val = r.value || '';
        if (val.includes(XML_NS) || val.includes(XML_NS_V1)) toDelete.push(i);
      });
      // Delete in reverse order so indices stay valid
      for (let i = toDelete.length - 1; i >= 0; i--) {
        parts.items[toDelete[i]].delete();
      }
      parts.add(xml);
      await ctx.sync();
      showStatus(msg || 'Saved.', 'success');
    });
  } catch (e) {
    showStatus('Save failed: ' + e.message, 'error');
  }
}

// ── XML v2 — Build ─────────────────────────────────────────────────────────────
function buildXmlV2() {
  const els = variables.map(function (v) {
    const base = 'id="' + x(v.id) + '" kind="' + x(v.kind) + '" name="' + x(v.name) + '" group="' + x(v.group || '') + '"';
    if (v.kind === 'scalar') {
      return '<variable ' + base + ' type="' + x(v.type || 'STRING') + '" defaultValue="' + x(v.defaultValue || '') + '"/>';
    }
    if (v.kind === 'computed') {
      return '<variable ' + base + ' formula="' + x(v.formula || '') + '"/>';
    }
    if (v.kind === 'system') {
      return '<variable ' + base + ' systemType="' + x(v.systemType || '') + '"/>';
    }
    if (v.kind === 'table') {
      const cols = (v.columns || []).map(function (c, i) {
        return '<col name="' + x(c) + '" type="' + x((v.columnTypes || [])[i] || 'STRING') + '"/>';
      }).join('');
      const rows = (v.rows || []).map(function (row) {
        return '<row>' + (row || []).map(function (cell) { return '<cell>' + x(cell) + '</cell>'; }).join('') + '</row>';
      }).join('');
      return '<variable ' + base + '><columns>' + cols + '</columns><rows>' + rows + '</rows></variable>';
    }
    return '';
  }).filter(Boolean).join('\n  ');

  return '<?xml version="1.0" encoding="UTF-8"?>\n' +
         '<livedocVariables xmlns="' + XML_NS + '" version="2">\n  ' +
         els + '\n</livedocVariables>';
}

// ── XML v2 — Parse ─────────────────────────────────────────────────────────────
function parseXmlV2(xmlStr) {
  const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  if (doc.querySelector('parsererror')) throw new Error('Malformed XML');
  variables = [];
  Array.from(doc.getElementsByTagName('variable')).forEach(function (el) {
    const kind = el.getAttribute('kind') || 'scalar';
    const base = {
      id:    el.getAttribute('id')    || uid(),
      kind:  kind,
      name:  el.getAttribute('name')  || '',
      group: el.getAttribute('group') || '',
    };
    if (kind === 'scalar') {
      variables.push(Object.assign({}, base, {
        type:         el.getAttribute('type')         || 'STRING',
        defaultValue: el.getAttribute('defaultValue') || '',
      }));
    } else if (kind === 'computed') {
      variables.push(Object.assign({}, base, {
        formula: el.getAttribute('formula') || '',
      }));
    } else if (kind === 'system') {
      const systemType = el.getAttribute('systemType') || '';
      const preset     = SYSTEM_PRESETS[systemType] || {};
      variables.push(Object.assign({}, base, {
        systemType:  systemType,
        columns:     preset.columns     || [],
        columnTypes: preset.columnTypes || [],
      }));
    } else if (kind === 'table') {
      const cols       = Array.from(el.getElementsByTagName('col'));
      const columns    = cols.map(function (c) { return c.getAttribute('name') || ''; });
      const columnTypes= cols.map(function (c) { return c.getAttribute('type') || 'STRING'; });
      const rows       = Array.from(el.getElementsByTagName('row')).map(function (row) {
        return Array.from(row.getElementsByTagName('cell')).map(function (c) { return c.textContent; });
      });
      variables.push(Object.assign({}, base, { columns: columns, columnTypes: columnTypes, rows: rows }));
    }
  });
}

// ── XML v1 — Migrate ───────────────────────────────────────────────────────────
function parseXmlV1(xmlStr) {
  const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  if (doc.querySelector('parsererror')) throw new Error('Malformed XML');
  variables = [];
  Array.from(doc.getElementsByTagName('variable')).forEach(function (el) {
    variables.push({
      id:           uid(),
      kind:         'scalar',
      name:         el.getAttribute('name')         || '',
      type:         el.getAttribute('type')         || 'STRING',
      defaultValue: el.getAttribute('defaultValue') || '',
      group:        el.getAttribute('group')        || '',
    });
  });
  Array.from(doc.getElementsByTagName('tableVariable')).forEach(function (el) {
    const columns = Array.from(el.getElementsByTagName('col')).map(function (c) { return c.textContent; });
    const rows    = Array.from(el.getElementsByTagName('row')).map(function (row) {
      return Array.from(row.getElementsByTagName('cell')).map(function (c) { return c.textContent; });
    });
    variables.push({
      id:          uid(),
      kind:        'table',
      name:        el.getAttribute('name')  || '',
      group:       el.getAttribute('group') || '',
      columns:     columns,
      columnTypes: columns.map(function () { return 'STRING'; }),
      rows:        rows,
    });
  });
}

// ── Insert a raw {{token}} string at the current cursor position ──────────────
function insertToken(tokenName) {
  var token = '{{' + tokenName + '}}';
  Office.context.document.setSelectedDataAsync(
    token,
    { coercionType: Office.CoercionType.Text },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showStatus('Inserted: ' + token, 'success');
      } else {
        showStatus('Click inside a text box or table cell first, then insert.', 'error');
      }
    }
  );
}

// ── Insert token at cursor ─────────────────────────────────────────────────────
function insertVariable(id) {
  const v = variables.find(function (v) { return v.id === id; });
  if (!v) return;

  if (v.kind === 'table' || (v.kind === 'system' && (v.columns || []).length > 0)) {
    insertTableShape(v);
  } else {
    const token = '{{' + v.name + '}}';
    Office.context.document.setSelectedDataAsync(
      token,
      { coercionType: Office.CoercionType.Text },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          showStatus('Inserted: ' + token, 'success');
        } else {
          showStatus('Click inside a text box on the slide first, then insert.', 'error');
        }
      }
    );
  }
}

// ── Insert table shape on current slide ───────────────────────────────────────
async function insertTableShape(v) {
  const columns  = v.columns  || [];
  const rows     = v.rows     || [];
  const colCount = columns.length;
  if (colCount === 0) { showStatus('Table has no columns defined.', 'error'); return; }

  const rowCount  = rows.length + 1; // +1 for header
  const tblWidth  = Math.min(680, Math.max(280, colCount * 130));
  const tblHeight = Math.max(40, rowCount * 36);

  try {
    await PowerPoint.run(async function (context) {
      let slide;
      try {
        const sel = context.presentation.getSelectedSlides();
        sel.load('items');
        await context.sync();
        slide = sel.items[0];
      } catch (_) {
        slide = context.presentation.slides.getItemAt(0);
      }

      const shape = slide.shapes.addTable(rowCount, colCount, {
        left: 50, top: 120, width: tblWidth, height: tblHeight,
      });
      await context.sync();

      const tbl = shape.table;
      // Header row
      for (let c = 0; c < colCount; c++) {
        tbl.rows.getItemAt(0).cells.getItemAt(c).text = columns[c] || '';
      }
      // Data rows (empty for system vars — filled at runtime)
      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < colCount; c++) {
          tbl.rows.getItemAt(r + 1).cells.getItemAt(c).text = (rows[r] || [])[c] || '';
        }
      }
      await context.sync();
      showStatus('Inserted table "' + v.name + '" (' + rowCount + ' rows \xd7 ' + colCount + ' cols).', 'success');
    });
  } catch (err) {
    showStatus('Table insert failed: ' + err.message, 'error');
  }
}

// ── Delete ─────────────────────────────────────────────────────────────────────
function deleteVariable(id) {
  const v = variables.find(function (v) { return v.id === id; });
  if (!v) return;
  if (!confirm('Delete "' + v.name + '"?')) return;
  variables = variables.filter(function (v) { return v.id !== id; });
  expandedIds.delete(id);
  if (editId === id) closeDrawer();
  saveToCustomXml('"' + v.name + '" deleted.');
  renderTree();
}

// ── Edit drawer ────────────────────────────────────────────────────────────────
function openEditDrawer(id) {
  const v = variables.find(function (v) { return v.id === id; });
  if (!v || v.kind === 'system') return;
  editMode = 'edit';
  editKind = v.kind;
  editId   = id;
  buildDrawerForm(v.kind, v);
  openDrawer('Edit ' + kindLabel(v.kind) + ': ' + v.name);
}

function buildDrawerForm(kind, data) {
  const form = document.getElementById('drawerForm');
  let html   = '';

  if (kind === 'scalar') {
    const typeOpts = ['STRING', 'NUMBER', 'DATE', 'BOOLEAN'].map(function (t) {
      return '<option value="' + t + '"' + (data && data.type === t ? ' selected' : '') + '>' + t + '</option>';
    }).join('');
    html = '<div class="form-row"><label>Name</label>' +
           '<input id="df-name" class="field" spellcheck="false" autocomplete="off" value="' + (data ? h(data.name) : '') + '"/></div>' +
           '<div class="form-row"><label>Type</label>' +
           '<select id="df-type" class="field">' + typeOpts + '</select></div>' +
           '<div class="form-row"><label>Default Value</label>' +
           '<input id="df-value" class="field" value="' + (data ? h(data.defaultValue || '') : '') + '"/></div>' +
           '<div class="form-row"><label>Group</label>' +
           '<input id="df-group" class="field" placeholder="e.g. Client Info" value="' + (data ? h(data.group || '') : '') + '"/></div>';

  } else if (kind === 'table') {
    const csvVal = data ? buildCsv(data) : '';
    html = '<div class="form-row"><label>Name</label>' +
           '<input id="df-name" class="field" spellcheck="false" autocomplete="off" value="' + (data ? h(data.name) : '') + '"/></div>' +
           '<div class="form-row"><label>Group</label>' +
           '<input id="df-group" class="field" placeholder="e.g. Financial" value="' + (data ? h(data.group || '') : '') + '"/></div>' +
           '<div class="form-row form-col"><label>CSV <span class="opt">first row = column headers</span></label>' +
           '<textarea id="df-csv" class="field textarea" rows="4">' + h(csvVal) + '</textarea></div>';

  } else if (kind === 'computed') {
    html = '<div class="form-row"><label>Name</label>' +
           '<input id="df-name" class="field" spellcheck="false" autocomplete="off" value="' + (data ? h(data.name) : '') + '"/></div>' +
           '<div class="form-row"><label>Formula</label>' +
           '<input id="df-formula" class="field mono" placeholder="=SUM(A1:A10)" value="' + (data ? h(data.formula || '') : '') + '"/></div>' +
           '<div class="form-row"><label>Group</label>' +
           '<input id="df-group" class="field" value="' + (data ? h(data.group || 'Computed') : 'Computed') + '"/></div>' +
           '<p class="form-hint">Evaluated by formulajs (Excel-compatible). Use variable names directly — e.g. =UPPER(ClientName) or =SUM(T1_Revenue).</p>';

  } else if (kind === 'system') {
    const presetOpts = Object.entries(SYSTEM_PRESETS).map(function (kv) {
      return '<option value="' + kv[0] + '">' + kv[0] + ' — ' + kv[1].description + '</option>';
    }).join('');
    html = '<div class="form-row"><label>Preset</label>' +
           '<select id="df-preset" class="field">' + presetOpts + '</select></div>' +
           '<p class="form-hint">System variables are predefined and evaluated at runtime by the engine. They are read-only in the panel.</p>';
  }

  form.innerHTML = html;
}

function buildCsv(tv) {
  const lines = [(tv.columns || []).join(',')];
  (tv.rows || []).forEach(function (row) { lines.push(row.join(',')); });
  return lines.join('\n');
}

function openDrawer(title) {
  document.getElementById('drawerTitle').textContent = title;
  document.getElementById('editDrawer').classList.remove('collapsed');
}

function closeDrawer() {
  document.getElementById('editDrawer').classList.add('collapsed');
  editId = null;
}

function saveDrawer() {
  if (editMode === 'dynamic')    { saveDynamicTableConfig();  return; }
  if (editMode === 'dynimage')   { saveDynamicImageConfig();  return; }
  if      (editKind === 'scalar')   saveScalar();
  else if (editKind === 'table')    saveTable();
  else if (editKind === 'computed') saveComputed();
  else if (editKind === 'system')   saveSystem();
}

function saveScalar() {
  const name  = (document.getElementById('df-name')  || {}).value || '';
  const type  = (document.getElementById('df-type')  || {}).value || 'STRING';
  const value = (document.getElementById('df-value') || {}).value || '';
  const group = (document.getElementById('df-group') || {}).value || '';

  if (!validateName(name.trim())) return;
  if (nameConflict(name.trim())) return;

  upsertVariable({
    id: editId || uid(), kind: 'scalar',
    name: name.trim(), type: type, defaultValue: value.trim(), group: group.trim(),
  });
  closeDrawer();
  saveToCustomXml('"' + name.trim() + '" ' + (editMode === 'edit' ? 'updated.' : 'added.'));
  renderTree();
}

function saveTable() {
  const name  = (document.getElementById('df-name')  || {}).value || '';
  const group = (document.getElementById('df-group') || {}).value || '';
  const csv   = (document.getElementById('df-csv')   || {}).value || '';

  if (!validateName(name.trim())) return;
  if (nameConflict(name.trim())) return;
  if (!csv.trim()) { showStatus('CSV data is required.', 'error'); return; }

  const lines      = csv.trim().split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
  const columns    = lines[0].split(',').map(function (c) { return c.trim(); });
  const columnTypes= columns.map(function () { return 'STRING'; });
  const rows       = lines.slice(1).map(function (line) {
    const cells = line.split(',').map(function (c) { return c.trim(); });
    while (cells.length < columns.length) cells.push('');
    return cells.slice(0, columns.length);
  });

  // Preserve existing column types when editing
  if (editMode === 'edit' && editId) {
    const existing = variables.find(function (v) { return v.id === editId; });
    if (existing && existing.columnTypes) {
      columns.forEach(function (col, i) {
        const prevIdx = (existing.columns || []).indexOf(col);
        if (prevIdx >= 0) columnTypes[i] = existing.columnTypes[prevIdx];
      });
    }
  }

  upsertVariable({
    id: editId || uid(), kind: 'table',
    name: name.trim(), group: group.trim(), columns: columns, columnTypes: columnTypes, rows: rows,
  });
  closeDrawer();
  saveToCustomXml('"' + name.trim() + '" ' + (editMode === 'edit' ? 'updated.' : 'added.'));
  renderTree();
}

function saveComputed() {
  const name    = (document.getElementById('df-name')    || {}).value || '';
  const formula = (document.getElementById('df-formula') || {}).value || '';
  const group   = (document.getElementById('df-group')   || {}).value || '';

  if (!validateName(name.trim())) return;
  if (nameConflict(name.trim())) return;

  upsertVariable({
    id: editId || uid(), kind: 'computed',
    name: name.trim(), formula: formula.trim(), group: group.trim() || 'Computed',
  });
  closeDrawer();
  saveToCustomXml('"' + name.trim() + '" ' + (editMode === 'edit' ? 'updated.' : 'added.'));
  renderTree();
}

function saveSystem() {
  const preset = (document.getElementById('df-preset') || {}).value;
  if (!preset || !SYSTEM_PRESETS[preset]) { showStatus('Please select a preset.', 'error'); return; }
  if (variables.find(function (v) { return v.kind === 'system' && v.systemType === preset; })) {
    showStatus('"' + preset + '" is already in the panel.', 'error'); return;
  }
  const p = SYSTEM_PRESETS[preset];
  upsertVariable({
    id: uid(), kind: 'system', name: preset, systemType: preset, group: 'System',
    columns: p.columns, columnTypes: p.columnTypes,
  });
  closeDrawer();
  saveToCustomXml('System variable "' + preset + '" added.');
  renderTree();
}

function upsertVariable(entry) {
  if (editMode === 'edit' && editId) {
    const idx = variables.findIndex(function (v) { return v.id === editId; });
    if (idx >= 0) { variables[idx] = entry; return; }
  }
  variables.push(entry);
}

function validateName(name) {
  if (!name) { showStatus('Name is required.', 'error'); return false; }
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    showStatus('Name must start with a letter or underscore, and contain only A–Z, 0–9, _.', 'error');
    return false;
  }
  return true;
}

function nameConflict(name) {
  const conflict = variables.find(function (v) {
    return v.name === name && v.id !== editId;
  });
  if (conflict) { showStatus('"' + name + '" already exists.', 'error'); return true; }
  return false;
}

// ── Tree rendering ─────────────────────────────────────────────────────────────
function renderTree() {
  const container = document.getElementById('varTree');
  document.getElementById('varCount').textContent = variables.length;

  const filtered = searchQuery
    ? variables.filter(function (v) {
        return v.name.toLowerCase().includes(searchQuery) ||
               (v.group || '').toLowerCase().includes(searchQuery);
      })
    : variables;

  if (!filtered.length) {
    container.innerHTML = variables.length === 0
      ? '<div class="tree-empty"><div class="tree-empty-icon">&#128203;</div>No variables yet — click <b>+ Add</b> to start.</div>'
      : '<div class="tree-empty"><div class="tree-empty-icon">&#128269;</div>No matches for &ldquo;' + h(searchQuery) + '&rdquo;.</div>';
    return;
  }

  // Build group map (preserving insertion order, ungrouped = '' key)
  const grouped = new Map();
  filtered.forEach(function (v) {
    const g = v.group || '';
    if (!grouped.has(g)) grouped.set(g, []);
    grouped.get(g).push(v);
  });

  let html = '';

  // Ungrouped first
  (grouped.get('') || []).forEach(function (v) { html += renderVarRow(v, false); });

  // Named groups
  grouped.forEach(function (items, group) {
    if (!group) return;
    html += '<div class="tree-group">' +
            '<div class="tree-group-header">' +
            '<span class="folder-toggle">&#9660;</span>' +
            '<span class="folder-icon">&#128193;</span>' +
            '<span class="group-name">' + h(group) + '</span>' +
            '<span class="group-count">' + items.length + '</span>' +
            '</div>' +
            '<div class="tree-group-body">' +
            items.map(function (v) { return renderVarRow(v, true); }).join('') +
            '</div></div>';
  });

  container.innerHTML = html;
  wireGroupToggles(container);
  wireExpandToggles(container);
}

function renderVarRow(v, inGroup) {
  const inGroupClass = inGroup ? ' in-group' : '';
  const isSystem     = v.kind === 'system';
  const hasChildren  = (v.kind === 'table' || v.kind === 'system') && (v.columns || []).length > 0;
  const isExpanded   = expandedIds.has(v.id);

  // ── Badge ──────────────────────────────
  let badge = '';
  if (v.kind === 'scalar') {
    const label = TYPE_LABELS[v.type || 'STRING'] || v.type;
    badge = '<span class="type-badge badge-' + h(v.type || 'STRING') + '" title="' + h(v.type || 'STRING') + '">' + label + '</span>';
  } else if (v.kind === 'table') {
    badge = '<span class="type-badge badge-TABLE" title="Table Variable">TBL</span>';
  } else if (v.kind === 'computed') {
    badge = '<span class="type-badge badge-COMPUTED" title="Computed Variable">fx</span>';
  } else if (v.kind === 'system') {
    badge = '<span class="type-badge badge-SYSTEM" title="System Variable">SYS</span>';
  }

  // ── Subtitle ───────────────────────────
  let subtitle = '';
  if (v.kind === 'scalar') {
    subtitle = v.defaultValue
      ? h(v.defaultValue)
      : '<em class="no-val">no default</em>';
  } else if (v.kind === 'table') {
    subtitle = (v.columns || []).length + ' cols \xd7 ' + (v.rows || []).length + ' rows';
  } else if (v.kind === 'computed') {
    subtitle = v.formula
      ? '<span class="formula-preview">' + h(v.formula) + '</span>'
      : '<em class="no-val">no formula</em>';
  } else if (v.kind === 'system') {
    subtitle = v.systemType || '';
  }

  // ── Expand toggle / spacer ─────────────
  const toggleHtml = hasChildren
    ? '<span class="row-expand-toggle" data-id="' + j(v.id) + '">' + (isExpanded ? '&#9660;' : '&#9654;') + '</span>'
    : '<span class="row-expand-spacer"></span>';

  // ── Action buttons ─────────────────────
  const insertTitle = hasChildren ? 'Insert table on current slide' : 'Insert token at cursor';
  const actInsert = '<button class="act-btn act-insert" onclick="insertVariable(\'' + j(v.id) + '\')" title="' + insertTitle + '">→</button>';
  const actEdit   = !isSystem
    ? '<button class="act-btn act-edit"   onclick="openEditDrawer(\'' + j(v.id) + '\')" title="Edit">✎</button>' : '';
  const actDelete = !isSystem
    ? '<button class="act-btn act-delete" onclick="deleteVariable(\'' + j(v.id) + '\')"  title="Delete">✕</button>' : '';

  let rowHtml = '<div class="tree-var-row' + inGroupClass + '">' +
                toggleHtml + badge +
                '<div class="var-info">' +
                  '<div class="var-name">' + h(v.name) + '</div>' +
                  '<div class="var-value">' + subtitle + '</div>' +
                '</div>' +
                '<div class="tree-actions">' + actInsert + actEdit + actDelete + '</div>' +
                '</div>';

  // ── Expanded children (column names) ───
  if (hasChildren && isExpanded) {
    (v.columns || []).forEach(function (col, i) {
      const ct   = (v.columnTypes || [])[i] || 'STRING';
      const clbl = TYPE_LABELS[ct] || ct.slice(0, 3);
      rowHtml += '<div class="tree-var-row tree-child-row' + inGroupClass + '">' +
                 '<span class="row-expand-spacer"></span>' +
                 '<span class="row-expand-spacer"></span>' +
                 '<span class="type-badge badge-' + h(ct) + '" title="' + h(ct) + '">' + clbl + '</span>' +
                 '<div class="var-info"><div class="var-name">' + h(col) + '</div></div>' +
                 '<div class="tree-actions">' +
                 '<button class="act-btn act-insert" onclick="insertToken(\'' + j(col) + '\')" title="Insert {{' + h(col) + '}} at cursor">→</button>' +
                 '</div>' +
                 '</div>';
    });
  }

  return rowHtml;
}

function wireGroupToggles(container) {
  container.querySelectorAll('.tree-group-header').forEach(function (header) {
    header.addEventListener('click', function () {
      const body     = header.nextElementSibling;
      const toggle   = header.querySelector('.folder-toggle');
      const collapsed= body.classList.toggle('collapsed');
      toggle.innerHTML = collapsed ? '&#9654;' : '&#9660;';
    });
  });
}

function wireExpandToggles(container) {
  container.querySelectorAll('.row-expand-toggle').forEach(function (btn) {
    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      const id = btn.dataset.id;
      if (expandedIds.has(id)) { expandedIds.delete(id); }
      else                     { expandedIds.add(id); }
      renderTree();
    });
  });
}

// ── Computed variable evaluation (formulajs) ──────────────────────────────────
//
// Builds a flat context from current variables so formulajs functions can
// reference them by name inside the formula string.
//
// Scalar   → bare value  (ClientName = "Acme",  Revenue = 5000)
// Table    → column arrays (T1_A = ["Alice","Bob"],  T1_B = [100,200])
//            plus the full table as array-of-rows  (T1 = [["Alice",100],…])
//
// NOTE: evaluation uses the Function() constructor so formulas run as JS.
// This is acceptable for an internal POC; for production, use the LDS engine.

function buildFormulaContext() {
  var ctx = {};
  variables.forEach(function (v) {
    if (v.kind === 'scalar') {
      var val = v.defaultValue || '';
      if      (v.type === 'NUMBER')  val = parseFloat(val)  || 0;
      else if (v.type === 'BOOLEAN') val = (val.toLowerCase() === 'true');
      ctx[v.name] = val;
    } else if (v.kind === 'table') {
      // Full table: array of row arrays
      ctx[v.name] = v.rows || [];
      // Individual column arrays: VarName_ColName
      (v.columns || []).forEach(function (col, i) {
        var colKey = v.name + '_' + col;
        ctx[colKey] = (v.rows || []).map(function (row) {
          var cell = (row || [])[i];
          return cell === undefined ? '' : cell;
        });
      });
    }
    // computed/system: not included in context
  });
  return ctx;
}

// Evaluate a formula string like "=UPPER(ClientName)" or "=SUM(T1_Revenue)".
// Returns the result as a string, or an error token like "#ERR: ...".
function evaluateFormula(formula) {
  if (typeof formulajs === 'undefined') {
    return '[formulajs not loaded]';
  }
  var expr = (formula || '').trim();
  if (expr.charAt(0) === '=') expr = expr.slice(1);
  if (!expr) return '';

  var ctx = buildFormulaContext();

  try {
    // Spread formulajs functions + variable context into the function scope
    var fnKeys = Object.keys(formulajs);
    var fnVals = fnKeys.map(function (k) { return formulajs[k]; });
    var ctxKeys = Object.keys(ctx);
    var ctxVals = ctxKeys.map(function (k) { return ctx[k]; });

    var allKeys = fnKeys.concat(ctxKeys);
    var allVals = fnVals.concat(ctxVals);

    // Build and call the evaluator function
    var fn     = new Function(allKeys, '"use strict"; return (' + expr + ');');
    var result = fn.apply(null, allVals);

    if (result === null || result === undefined) return '';
    if (typeof result === 'object') return JSON.stringify(result);
    return String(result);
  } catch (e) {
    return '#ERR: ' + String(e.message).slice(0, 60);
  }
}

// ── Selection tracking ─────────────────────────────────────────────────────────

function initSelectionHandler() {
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSlideSelectionChanged,
    function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn('SelectionChanged handler registration failed:', result.error && result.error.message);
      }
    }
  );
}

async function onSlideSelectionChanged() {
  try {
    var newSlideId  = null;
    var foundTable  = null;
    var foundShape  = null; // any non-indicator shape
    var dynConfig   = null;
    var imgConfig   = null;

    await PowerPoint.run(async function (context) {
      // Detect current slide ID so we can tell when the user switches slides
      try {
        var selSlides = context.presentation.getSelectedSlides();
        selSlides.load('items/id');
        await context.sync();
        newSlideId = selSlides.items.length ? selSlides.items[0].id : null;
      } catch (_) {}

      // Detect selected shapes
      try {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load('items/name,items/type');
        await context.sync();
        foundTable = shapes.items.find(function (s) {
          return s.type === 'Table' ||
                 (typeof PowerPoint !== 'undefined' &&
                  typeof PowerPoint.ShapeType !== 'undefined' &&
                  s.type === PowerPoint.ShapeType.table);
        }) || null;
        foundShape = shapes.items.find(function (s) {
          return !s.name.startsWith('__LIVEDOC_IND_');
        }) || null;
      } catch (_) {}

      // Load dynamic table config if table drawer is open
      if (editMode === 'dynamic' && foundTable) {
        try {
          foundTable.tags.load('items/key,items/value');
          await context.sync();
          var dynTag = foundTable.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_TABLE'; });
          dynConfig = dynTag ? JSON.parse(dynTag.value) : null;
        } catch (_) {}
      }

      // Load dynamic image config if image drawer is open
      if (editMode === 'dynimage' && foundShape) {
        try {
          foundShape.tags.load('items/key,items/value');
          await context.sync();
          var imgTag = foundShape.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_IMAGE'; });
          imgConfig = imgTag ? JSON.parse(imgTag.value) : null;
        } catch (_) {}
      }
    });

    var slideChanged  = (newSlideId !== null && newSlideId !== currentSlideId);
    if (newSlideId !== null) currentSlideId = newSlideId;
    selectedTableName = foundTable ? foundTable.name : null;
    selectedShapeName = foundShape ? foundShape.name : null;

    if (editMode === 'dynamic') {
      if (slideChanged) {
        await refreshDynamicDrawer();
      } else if (selectedTableName) {
        syncDynamicDrawerDropdown(selectedTableName);
        syncDynamicDrawerFromConfig(dynConfig);
      }
    } else if (editMode === 'dynimage') {
      if (slideChanged) {
        await refreshDynamicImageDrawer();
      } else if (selectedShapeName) {
        syncDynamicImageDrawerDropdown(selectedShapeName);
        syncDynamicImageDrawerFromConfig(imgConfig);
      }
    }
  } catch (_) {
    selectedTableName = null;
    selectedShapeName = null;
  }
}

// Reload the Dynamic drawer form for the current slide (called when the slide changes).
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

// Sync the Variable dropdown and Repeat Rows fields to an existing config (or reset to defaults).
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
  } else {
    varSel.selectedIndex = 0;
    fromEl.value = 2;
    toEl.value   = 2;
  }
}

// ── Dynamic Table Settings ─────────────────────────────────────────────────────
//
// A PPT table shape is "tagged" as dynamic by storing a JSON config in
// shape.tags under the key LIVEDOC_DYN_TABLE.  At preview time we:
//   1. Scan all slides via Office.js API to collect those configs.
//   2. In the JSZip pass, find the matching <p:graphicFrame> by shape name,
//      clone the template <a:tr> rows for every data record, then drop the
//      original template rows.

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
      // ShapeType.table = 'Table' in Office.js string enum
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

  form.innerHTML =
    '<div class="form-row"><label>Slide Table</label>' +
    '<select id="df-dyn-shape" class="field">' + shapeOpts + '</select></div>' +
    '<div class="form-row"><label>Variable</label>' +
    '<select id="df-dyn-var" class="field">' + varOpts + '</select></div>' +
    '<div class="form-row"><label>Repeat Rows</label>' +
    '<div class="row-range-wrap">' +
    '<span class="range-label">From</span>' +
    '<input id="df-dyn-from" class="field field-sm" type="number" min="1" value="2"/>' +
    '<span class="range-label">To</span>' +
    '<input id="df-dyn-to"   class="field field-sm" type="number" min="1" value="2"/>' +
    '<span class="range-hint">(1-indexed row numbers)</span>' +
    '</div></div>' +
    '<p class="form-hint">Rows <b>From</b>–<b>To</b> in the slide table are the repeating template. ' +
    'At preview they are expanded once per data record, then the template is removed.</p>';
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

  var config = { variableName: varName, fromRow: fromRow, toRow: toRow };

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

// ── Dynamic Image Settings ─────────────────────────────────────────────────────
//
// A slide shape is tagged as a dynamic image placeholder by storing a JSON config
// in shape.tags under LIVEDOC_DYN_IMAGE: { fitMode, sourceUrl }.
// At preview time the placeholder <p:sp> is replaced by a <p:pic> element with
// the image embedded in the zip, fitted to the original shape's EMU bounds.
// A small companion indicator shape (__LIVEDOC_IND_<name>) is added to the slide
// to show the 📷 badge; these are stripped from all slides during preview.

async function openDynamicImageDrawer() {
  editMode = 'dynimage';
  pendingImageDataUrl = null;
  openDrawer('Dynamic Image Settings');
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
  buildDynamicImageDrawerForm(shapes);
}

async function loadCurrentSlideAllShapes() {
  var shapes = [];
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
      if (!s.name.startsWith('__LIVEDOC_IND_')) {
        shapes.push({ name: s.name, type: s.type });
      }
    });
  });
  return shapes;
}

function buildDynamicImageDrawerForm(shapes) {
  var form = document.getElementById('drawerForm');
  if (!shapes.length) {
    form.innerHTML = '<p class="form-hint">No shapes found on the current slide.<br>Add a shape to use as a placeholder first.</p>';
    return;
  }

  var shapeOpts = shapes.map(function (s) {
    var sel = (selectedShapeName && s.name === selectedShapeName) ? ' selected' : '';
    return '<option value="' + h(s.name) + '"' + sel + '>' + h(s.name) + ' (' + h(s.type) + ')</option>';
  }).join('');

  var fileStatusMsg = pendingImageDataUrl
    ? '<span style="color:#1a6b3a">✓ File loaded — ready to save</span>'
    : '';

  form.innerHTML =
    '<div class="form-row"><label>Placeholder Shape</label>' +
    '<select id="df-img-shape" class="field">' + shapeOpts + '</select></div>' +
    '<div class="form-row"><label>Fit Mode</label>' +
    '<select id="df-img-fit" class="field">' +
    '<option value="fitInside">Fit Inside (keep aspect ratio)</option>' +
    '<option value="fitWidth">Fit Width (keep aspect ratio)</option>' +
    '<option value="fitHeight">Fit Height (keep aspect ratio)</option>' +
    '<option value="stretch">Stretch (fill bounds, ignore aspect ratio)</option>' +
    '</select></div>' +
    '<div class="form-row form-col"><label>Image URL</label>' +
    '<input id="df-img-url" class="field" type="url" placeholder="https://example.com/photo.jpg"/></div>' +
    '<div class="form-row form-col"><label>— or upload local file —</label>' +
    '<input id="df-img-file" type="file" accept="image/*" class="field" onchange="onImageFileSelected(this)"/>' +
    '<div id="df-img-file-status" class="form-hint" style="margin-top:2px">' + fileStatusMsg + '</div>' +
    '</div>' +
    '<p class="form-hint">At preview the placeholder is replaced by the image, fitted to its bounds. ' +
    'A 📷 indicator is added to the slide to show the binding.</p>';
}

async function refreshDynamicImageDrawer() {
  var form = document.getElementById('drawerForm');
  if (!form) return;
  form.innerHTML = '<p class="form-hint" style="padding:8px 0">Loading shapes from current slide…</p>';
  try {
    var shapes = await loadCurrentSlideAllShapes();
    buildDynamicImageDrawerForm(shapes);
  } catch (e) {
    form.innerHTML = '<p class="form-hint" style="color:#c00">Could not load shapes: ' + h(e.message) + '</p>';
  }
}

function syncDynamicImageDrawerDropdown(shapeName) {
  var sel = document.getElementById('df-img-shape');
  if (!sel || !shapeName) return;
  for (var i = 0; i < sel.options.length; i++) {
    if (sel.options[i].value === shapeName) { sel.selectedIndex = i; break; }
  }
}

function syncDynamicImageDrawerFromConfig(config) {
  var fitSel = document.getElementById('df-img-fit');
  var urlEl  = document.getElementById('df-img-url');
  if (!fitSel || !urlEl) return;
  if (config) {
    for (var i = 0; i < fitSel.options.length; i++) {
      if (fitSel.options[i].value === config.fitMode) { fitSel.selectedIndex = i; break; }
    }
    // Don't populate URL field with data: URLs (base64 is too large to show)
    urlEl.value = (config.sourceUrl && !config.sourceUrl.startsWith('data:')) ? config.sourceUrl : '';
  }
}

async function saveDynamicImageConfig() {
  var shapeName = (document.getElementById('df-img-shape') || {}).value || '';
  var fitMode   = (document.getElementById('df-img-fit')   || {}).value || 'fitInside';
  var url       = ((document.getElementById('df-img-url')  || {}).value || '').trim();
  var fileInput = document.getElementById('df-img-file');

  if (!shapeName) { showStatus('Select a placeholder shape.', 'error'); return; }
  var hasFile = fileInput && fileInput.files && fileInput.files.length > 0;
  if (!url && !hasFile && !pendingImageDataUrl) {
    showStatus('Provide an image URL or upload a local file.', 'error'); return;
  }

  var sourceUrl = url;
  if (!sourceUrl && hasFile) {
    try { sourceUrl = await readFileAsDataUrl(fileInput.files[0]); }
    catch (e) { showStatus('File read failed: ' + e.message, 'error'); return; }
  }
  if (!sourceUrl && pendingImageDataUrl) {
    sourceUrl = pendingImageDataUrl;
  }
  pendingImageDataUrl = null; // consumed

  var config = { fitMode: fitMode, sourceUrl: sourceUrl };

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

      slide.shapes.load('items/name,items/left,items/top,items/width,items/height');
      await context.sync();

      var shape = slide.shapes.items.find(function (s) { return s.name === shapeName; });
      if (!shape) { showStatus('Shape "' + shapeName + '" not found on current slide.', 'error'); return; }

      // Capture bounds before deleting
      var left   = shape.left;
      var top    = shape.top;
      var width  = shape.width;
      var height = shape.height;

      // Remove the original shape and any stale indicator
      shape.delete();
      var staleInd = slide.shapes.items.find(function (s) { return s.name === '__LIVEDOC_IND_' + shapeName; });
      if (staleInd) staleInd.delete();
      await context.sync();

      // Replace with a visually distinct placeholder textbox at the same bounds
      var label = '📷  Dynamic Image\n' + shapeName;
      var ph = slide.shapes.addTextBox(label, { left: left, top: top, width: width, height: height });
      await context.sync();

      ph.name = shapeName;                        // keep original name so OOXML lookup still works
      ph.fill.setSolidColor('DCE8F8');            // light blue-gray — visually distinct from real content
      ph.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone; // keep original bounds
      ph.textFrame.textRange.font.size  = 14;
      ph.textFrame.textRange.font.color = '3A5A8A';
      ph.textFrame.textRange.font.bold  = false;
      try {
        // Center text if API available (requires 1.4+)
        ph.textFrame.textRange.paragraphFormat.horizontalAlignment =
          PowerPoint.ParagraphHorizontalAlignment.center;
        ph.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      } catch (_) {}

      ph.tags.add('LIVEDOC_DYN_IMAGE', JSON.stringify(config));
      await context.sync();

      closeDrawer();
      showStatus('"' + shapeName + '" placeholder ready (' + fitMode + '). Run Preview to replace with the real image.', 'success');
    });
  } catch (e) {
    showStatus('Failed to configure placeholder: ' + e.message, 'error');
  }
}

function readFileAsDataUrl(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload  = function (e) { resolve(e.target.result); };
    reader.onerror = function ()  { reject(new Error('File read failed')); };
    reader.readAsDataURL(file);
  });
}

// Called by the file input onchange — reads the file immediately into
// pendingImageDataUrl so the data survives if the drawer form is rebuilt.
function onImageFileSelected(input) {
  var statusEl = document.getElementById('df-img-file-status');
  if (!input.files || !input.files.length) {
    pendingImageDataUrl = null;
    if (statusEl) statusEl.innerHTML = '';
    return;
  }
  var file = input.files[0];
  if (statusEl) statusEl.innerHTML = '<span style="color:#888">Reading file…</span>';
  readFileAsDataUrl(file).then(function (dataUrl) {
    pendingImageDataUrl = dataUrl;
    if (statusEl) statusEl.innerHTML = '<span style="color:#1a6b3a">✓ Loaded: ' + h(file.name) + '</span>';
  }).catch(function () {
    pendingImageDataUrl = null;
    if (statusEl) statusEl.innerHTML = '<span style="color:#c00">File read failed.</span>';
  });
}

// Pre-scan all slides for dynamic table shape tags (called at preview start).
// Returns: { slideIndex(1-based): [{shapeName, variableName, fromRow, toRow}] }
async function scanDynamicTables() {
  var result = {};

  await PowerPoint.run(async function (context) {
    var slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    // Batch-load shape name+type for every slide
    slides.items.forEach(function (slide) {
      slide.shapes.load('items/name,items/type');
    });
    await context.sync();

    // Collect table shapes per slide
    var tableShapesBySlide = [];
    slides.items.forEach(function (slide, si) {
      var tShapes = slide.shapes.items.filter(function (s) {
        return s.type === 'Table' ||
               (typeof PowerPoint.ShapeType !== 'undefined' && s.type === PowerPoint.ShapeType.table);
      });
      if (tShapes.length) tableShapesBySlide.push({ slideIndex: si + 1, shapes: tShapes });
    });

    if (!tableShapesBySlide.length) return;

    // Batch-load tags for all table shapes
    tableShapesBySlide.forEach(function (entry) {
      entry.shapes.forEach(function (s) { s.tags.load('items/key,items/value'); });
    });
    await context.sync();

    // Extract LIVEDOC_DYN_TABLE configs
    tableShapesBySlide.forEach(function (entry) {
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
    var insertBefore = templateRows[0];   // new rows go before the first template row

    // For each data record: clone each template row and substitute {{ColName}}
    tv.rows.forEach(function (dataRow) {
      var rowMap = {};
      (tv.columns || []).forEach(function (col, i) { rowMap[col] = (dataRow || [])[i] || ''; });

      templateRows.forEach(function (templateRow) {
        var newRow = templateRow.cloneNode(true);

        // Replace {{ColName}} tokens inside every <a:t> text node
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

// ── Dynamic Image — scan & expand ─────────────────────────────────────────────
//
// scanDynamicImages(): identical pattern to scanDynamicTables() but reads
//   LIVEDOC_DYN_IMAGE tags from ALL shapes (not just tables).
// expandDynamicImagesInXml(): replaces the tagged <p:sp> with a <p:pic> that
//   embeds the image bytes directly in the PPTX zip, and strips all __LIVEDOC_IND_*
//   indicator shapes from the slide XML.

async function scanDynamicImages() {
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
        var imgTag = s.tags.items.find(function (t) { return t.key === 'LIVEDOC_DYN_IMAGE'; });
        if (!imgTag) return;
        try {
          var config = JSON.parse(imgTag.value);
          if (!result[entry.slideIndex]) result[entry.slideIndex] = [];
          result[entry.slideIndex].push(Object.assign({ shapeName: s.name }, config));
        } catch (_) {}
      });
    });
  });

  return result;
}

// Expand dynamic images in a single slide XML string.
// Also strips __LIVEDOC_IND_* indicator shapes from all slides (always called).
// Modifies zip in-place (adds ppt/media/* and updates rels).
// Returns the modified slide XML string.
async function expandDynamicImagesInXml(zip, xml, slideIndex, configs) {
  var NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

  var parser   = new DOMParser();
  var slideDoc = parser.parseFromString(xml, 'application/xml');
  if (slideDoc.getElementsByTagName('parseerror').length) return xml;

  // Always strip indicator shapes
  var spEls = Array.from(slideDoc.getElementsByTagNameNS(NS_P, 'sp'));
  spEls.forEach(function (sp) {
    var nvSpPr = sp.getElementsByTagNameNS(NS_P, 'nvSpPr')[0];
    if (!nvSpPr) return;
    var cNvPr = findChildByLocalName(nvSpPr, 'cNvPr');
    if (cNvPr && (cNvPr.getAttribute('name') || '').startsWith('__LIVEDOC_IND_')) {
      sp.parentNode.removeChild(sp);
    }
  });

  if (!configs || !configs.length) {
    var ser = new XMLSerializer();
    return ser.serializeToString(slideDoc).replace(/ xmlns=""/g, '');
  }

  // Load / create the slide rels XML
  var relsPath = 'ppt/slides/_rels/slide' + slideIndex + '.xml.rels';
  var relsXml  = zip.files[relsPath]
    ? await zip.files[relsPath].async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' + NS_REL + '"></Relationships>';

  var relsDoc = parser.parseFromString(relsXml, 'application/xml');

  var mediaSeq = 0; // per-slide counter to keep media names unique
  var changed  = false;

  for (var ci = 0; ci < configs.length; ci++) {
    var cfg = configs[ci];

    // Find the placeholder <p:sp> and its EMU bounds
    var bounds = findShapeBoundsInXml(slideDoc, cfg.shapeName, NS_P, NS_A);
    if (!bounds) {
      console.warn('[DynImage] Shape "' + cfg.shapeName + '" not found in slide XML.');
      continue;
    }

    // Fetch image bytes + mime type
    var imgData;
    try { imgData = await fetchImageData(cfg.sourceUrl); }
    catch (e) {
      console.warn('[DynImage] Could not fetch image for "' + cfg.shapeName + '":', e.message);
      continue;
    }

    // Get natural dimensions for aspect-ratio fit modes
    var dims = { w: null, h: null };
    if (cfg.fitMode !== 'stretch') {
      try { dims = await getImageDimensions(imgData.bytes, imgData.mimeType); } catch (_) {}
    }

    // Add image bytes to zip media folder
    var ext      = (imgData.mimeType.split('/')[1] || 'png').replace('jpeg', 'jpg');
    var mediaName = 'livedoc_dyn_s' + slideIndex + '_' + (++mediaSeq) + '.' + ext;
    zip.file('ppt/media/' + mediaName, imgData.bytes);

    // Ensure [Content_Types].xml has a Default entry for this extension.
    // Missing entries cause PowerPoint to prompt "needs repair" on open.
    await ensureContentType(zip, ext, imgData.mimeType);

    // Add relationship entry
    var rId = 'rIdLDDyn' + slideIndex + '_' + mediaSeq;
    var relEl = relsDoc.createElementNS(NS_REL, 'Relationship');
    relEl.setAttribute('Id',     rId);
    relEl.setAttribute('Type',   'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
    relEl.setAttribute('Target', '../media/' + mediaName);
    relsDoc.documentElement.appendChild(relEl);

    // Compute fillRect offsets for the chosen fit mode
    var fr = computeFillRect(dims.w, dims.h, bounds.cx, bounds.cy, cfg.fitMode);

    // Build <p:pic> and import it into the slide document
    var picXmlStr =
      '<p:pic' +
      ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<p:nvPicPr>' +
      '<p:cNvPr id="' + (parseInt(bounds.shapeId, 10) + 5000) + '" name="DynImg_' + h(cfg.shapeName) + '"/>' +
      '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>' +
      '<p:nvPr/>' +
      '</p:nvPicPr>' +
      '<p:blipFill>' +
      '<a:blip r:embed="' + rId + '"/>' +
      '<a:stretch><a:fillRect' +
        (fr.l ? ' l="' + fr.l + '"' : '') +
        (fr.t ? ' t="' + fr.t + '"' : '') +
        (fr.r ? ' r="' + fr.r + '"' : '') +
        (fr.b ? ' b="' + fr.b + '"' : '') +
      '/></a:stretch>' +
      '</p:blipFill>' +
      '<p:spPr>' +
      '<a:xfrm><a:off x="' + bounds.x + '" y="' + bounds.y + '"/>' +
      '<a:ext cx="' + bounds.cx + '" cy="' + bounds.cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '</p:spPr>' +
      '</p:pic>';

    var wrapDoc  = parser.parseFromString('<root>' + picXmlStr + '</root>', 'application/xml');
    var picEl    = wrapDoc.documentElement.firstElementChild;
    var imported = slideDoc.importNode(picEl, true);

    // Replace the placeholder shape with the pic
    bounds.spEl.parentNode.insertBefore(imported, bounds.spEl);
    bounds.spEl.parentNode.removeChild(bounds.spEl);
    changed = true;
  }

  if (changed) {
    var serializer = new XMLSerializer();
    zip.file(relsPath, serializer.serializeToString(relsDoc).replace(/ xmlns=""/g, ''));
  }

  var ser2 = new XMLSerializer();
  return ser2.serializeToString(slideDoc).replace(/ xmlns=""/g, '');
}

// ── Dynamic Image helpers ──────────────────────────────────────────────────────

function findChildByLocalName(parent, localName) {
  for (var i = 0; i < parent.childNodes.length; i++) {
    var n = parent.childNodes[i];
    if (n.nodeType === 1 && n.localName === localName) return n;
  }
  return null;
}

function findShapeBoundsInXml(slideDoc, shapeName, NS_P, NS_A) {
  // Search <p:sp> (regular shapes/textboxes) AND <p:pic> (image shapes).
  // Both use nvSpPr/nvPicPr > cNvPr[@name] and spPr > a:xfrm for position.
  var candidates = Array.from(slideDoc.getElementsByTagNameNS(NS_P, 'sp'))
    .concat(Array.from(slideDoc.getElementsByTagNameNS(NS_P, 'pic')));

  for (var i = 0; i < candidates.length; i++) {
    var el    = candidates[i];
    // nvSpPr for <p:sp>, nvPicPr for <p:pic>
    var nvPr  = el.getElementsByTagNameNS(NS_P, 'nvSpPr')[0] ||
                el.getElementsByTagNameNS(NS_P, 'nvPicPr')[0];
    if (!nvPr) continue;
    var cNvPr = findChildByLocalName(nvPr, 'cNvPr');
    if (!cNvPr || cNvPr.getAttribute('name') !== shapeName) continue;

    var spPr = el.getElementsByTagNameNS(NS_P, 'spPr')[0];
    if (!spPr) continue;
    var xfrm = spPr.getElementsByTagNameNS(NS_A, 'xfrm')[0];
    if (!xfrm) continue;
    var off  = xfrm.getElementsByTagNameNS(NS_A, 'off')[0];
    var ext  = xfrm.getElementsByTagNameNS(NS_A, 'ext')[0];
    if (!off || !ext) continue;

    return {
      spEl:    el,
      shapeId: cNvPr.getAttribute('id') || '99',
      x:   parseInt(off.getAttribute('x')  || '0', 10),
      y:   parseInt(off.getAttribute('y')  || '0', 10),
      cx:  parseInt(ext.getAttribute('cx') || '0', 10),
      cy:  parseInt(ext.getAttribute('cy') || '0', 10),
    };
  }
  return null;
}

// Returns fillRect offset values (1000ths-of-percent, 100%=100000) for the
// four edges, so the image is letterboxed/pillarboxed inside the shape bounds.
function computeFillRect(iw, ih, cx, cy, fitMode) {
  var l = 0, t = 0, r = 0, b = 0;
  if (fitMode === 'stretch' || !iw || !ih || !cx || !cy) return { l: l, t: t, r: r, b: b };

  var shapeR = cx / cy;
  var imageR = iw / ih;

  if (fitMode === 'fitInside') {
    if (imageR > shapeR) {
      // Image wider than shape → fit to width, letterbox top/bottom
      var scaledH = ih * cx / iw;
      var emptyV  = cy - scaledH;
      t = b = Math.max(0, Math.round(emptyV / 2 / cy * 100000));
    } else {
      // Image taller than shape → fit to height, pillarbox left/right
      var scaledW = iw * cy / ih;
      var emptyH  = cx - scaledW;
      l = r = Math.max(0, Math.round(emptyH / 2 / cx * 100000));
    }
  } else if (fitMode === 'fitWidth') {
    var scaledH2 = ih * cx / iw;
    if (scaledH2 < cy) {
      var emptyV2 = cy - scaledH2;
      t = b = Math.max(0, Math.round(emptyV2 / 2 / cy * 100000));
    }
  } else if (fitMode === 'fitHeight') {
    var scaledW2 = iw * cy / ih;
    if (scaledW2 < cx) {
      var emptyH2 = cx - scaledW2;
      l = r = Math.max(0, Math.round(emptyH2 / 2 / cx * 100000));
    }
  }

  return { l: l, t: t, r: r, b: b };
}

// Fetch image from URL or decode a data: URI. Returns { bytes: Uint8Array, mimeType }.
async function fetchImageData(sourceUrl) {
  if (sourceUrl.startsWith('data:')) {
    var m = sourceUrl.match(/^data:([^;]+);base64,(.+)$/);
    if (!m) throw new Error('Invalid data URL');
    var bin = atob(m[2]);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return { bytes: arr, mimeType: m[1] };
  }
  var resp = await fetch(sourceUrl);
  if (!resp.ok) throw new Error('HTTP ' + resp.status + ' fetching image');
  var mime = (resp.headers.get('content-type') || 'image/png').split(';')[0].trim();
  var ab   = await resp.arrayBuffer();
  return { bytes: new Uint8Array(ab), mimeType: mime };
}

// Register the image extension in [Content_Types].xml if not already present.
// Missing entries cause PowerPoint to show "needs repair" when opening the file.
async function ensureContentType(zip, ext, mimeType) {
  var ctPath = '[Content_Types].xml';
  if (!zip.files[ctPath]) return;
  var ct = await zip.files[ctPath].async('string');
  if (ct.includes('Extension="' + ext + '"')) return; // already registered

  // Normalise mime type (browsers may return 'image/jpeg' for jpg)
  var mime = mimeType || ('image/' + ext);
  if (ext === 'jpg') mime = 'image/jpeg';

  var entry = '<Default Extension="' + ext + '" ContentType="' + mime + '"/>';
  ct = ct.replace('</Types>', entry + '</Types>');
  zip.file(ctPath, ct);
}

// Load image bytes into an <img> element to get natural dimensions.
function getImageDimensions(bytes, mimeType) {
  return new Promise(function (resolve) {
    var blob = new Blob([bytes], { type: mimeType });
    var url  = URL.createObjectURL(blob);
    var img  = new Image();
    img.onload  = function () { URL.revokeObjectURL(url); resolve({ w: img.naturalWidth,  h: img.naturalHeight }); };
    img.onerror = function () { URL.revokeObjectURL(url); resolve({ w: null, h: null }); };
    img.src = url;
  });
}

// ── Preview — replace tokens and download PPTX ────────────────────────────────
//
// Flow: getFileAsync (sliced bytes) → JSZip → regex-replace {{Var}} in each
// slide XML → re-zip → browser download.
//
// Split-run handling: PowerPoint sometimes fragments {{VarName}} across
// multiple <a:r> runs (spell-check, paste, re-formatting). We first do a
// paragraph-level merge pass (DOM-based) before the regex pass so that tokens
// spread across runs are healed before replacement.

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
    var dynSlideCount   = Object.keys(dynamicTableMap).length;

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
      // (always called so indicators are removed even on slides with no images)
      var imgConfigs = dynamicImageMap[slideIndex] || [];
      xml = await expandDynamicImagesInXml(zip, xml, slideIndex, imgConfigs);
      if (imgConfigs.length) expandedImages += imgConfigs.length;

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
// a token that is split across run boundaries, merge the *minimal* adjacent
// runs needed to make the token contiguous, preserving the first run's rPr.

function healSplitTokenRuns(xml, replacements) {
  // Only operate on paragraphs whose full text contains an opening {{ not
  // matched in a single <a:t> element — quick pre-check to avoid DOM overhead.
  if (!xml.includes('{{')) return xml;

  // Use DOMParser to work with the paragraph structure
  var parser = new DOMParser();
  var doc;
  try {
    doc = parser.parseFromString(xml, 'application/xml');
  } catch (e) {
    return xml; // can't parse — fall through to regex pass
  }
  if (doc.querySelector('parsererror')) return xml;

  var changed = false;
  var paragraphs = doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p');

  for (var pi = 0; pi < paragraphs.length; pi++) {
    var para    = paragraphs[pi];
    var runs    = Array.from(para.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'r'));
    if (runs.length < 2) continue;

    // Concatenate all run texts
    var texts = runs.map(function (r) {
      var tEl = r.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't')[0];
      return tEl ? tEl.textContent : '';
    });
    var combined = texts.join('');

    // Check if combined text has a token that spans a run boundary
    var tokenRe = /\{\{(\w+)\}\}/g;
    var needsMerge = false;
    var m;
    while ((m = tokenRe.exec(combined)) !== null) {
      var varName = m[1];
      if (!Object.prototype.hasOwnProperty.call(replacements, varName)) continue;
      // Does this token appear inside a single run as-is?
      var foundInSingleRun = texts.some(function (t) { return t.includes(m[0]); });
      if (!foundInSingleRun) { needsMerge = true; break; }
    }
    if (!needsMerge) continue;

    // Merge all runs into the first run, keeping first run's rPr, clearing rest
    var firstRun = runs[0];
    var firstT   = firstRun.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't')[0];
    if (firstT) {
      firstT.textContent = combined;
      // Preserve xml:space="preserve" if combined text has leading/trailing spaces
      if (combined !== combined.trim()) {
        firstT.setAttribute('xml:space', 'preserve');
      }
    }
    // Remove all runs after the first
    for (var ri = 1; ri < runs.length; ri++) {
      runs[ri].parentNode.removeChild(runs[ri]);
    }
    changed = true;
  }

  if (!changed) return xml;

  // Serialize back
  var serializer = new XMLSerializer();
  var newXml = serializer.serializeToString(doc);
  // XMLSerializer adds xmlns declarations — strip any spurious xmlns="" added to inner elements
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
              slices[sliceResult.value.index] = sliceResult.value.data; // Uint8Array
              remaining--;
              if (remaining === 0) {
                file.closeAsync();
                // Concatenate slices into one Uint8Array
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

// ── XML value escape (for injecting replacement text into OOXML) ───────────────
function xmlEscapeVal(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Helpers ────────────────────────────────────────────────────────────────────
function kindLabel(kind) {
  var map = { scalar: 'Scalar Variable', table: 'Table Variable', computed: 'Computed Variable', system: 'System Variable' };
  return map[kind] || kind;
}

function uid() {
  return Math.random().toString(36).slice(2, 10) + Math.random().toString(36).slice(2, 10);
}

var _statusTimer = null;
function showStatus(msg, type) {
  type = type || 'info';
  var el = document.getElementById('status');
  el.textContent  = msg;
  el.className    = 'status-bar ' + type;
  el.style.display= 'block';
  clearTimeout(_statusTimer);
  if (type !== 'error') {
    _statusTimer = setTimeout(function () { el.style.display = 'none'; }, 4000);
  }
}

// XML-safe attribute escape
function x(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
                  .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}
// HTML-safe display escape
function h(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
// JS single-quote escape for onclick attributes
function j(s) {
  return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'");
}
