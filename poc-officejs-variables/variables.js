'use strict';

// ── Variable CRUD ──────────────────────────────────────────────────────────────

function deleteVariable(id) {
  const v = variables.find(function (v) { return v.id === id; });
  if (!v) return;

  variables = variables.filter(function (v) { return v.id !== id; });
  expandedIds.delete(id);
  if (editId === id) closeDrawer();
  saveToCustomXml('"' + v.name + '" deleted.');
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

// ── Drawer open/close ──────────────────────────────────────────────────────────

function openAddDrawer(kind) {
  document.getElementById('addMenu').classList.add('hidden');
  editMode = 'add';
  editKind = kind;
  editId   = null;
  buildDrawerForm(kind, null);
  openDrawer('Add ' + kindLabel(kind));
}

function openEditDrawer(id) {
  const v = variables.find(function (v) { return v.id === id; });
  if (!v || v.kind === 'system') return;
  editMode = 'edit';
  editKind = v.kind;
  editId   = id;
  buildDrawerForm(v.kind, v);
  openDrawer('Edit ' + kindLabel(v.kind) + ': ' + v.name);
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
  if (editMode === 'dynchart')   { saveDynamicChartConfig();  return; }
  if      (editKind === 'scalar')   saveScalar();
  else if (editKind === 'table')    saveTable();
  else if (editKind === 'computed') saveComputed();
  else if (editKind === 'system')   saveSystem();
}

// ── Drawer forms ───────────────────────────────────────────────────────────────

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
           '<input id="df-group" class="field" value="Adhoc" readonly/></div>';

  } else if (kind === 'table') {
    const csvVal = data ? buildCsv(data) : '';
    html = '<div class="form-row"><label>Name</label>' +
           '<input id="df-name" class="field" spellcheck="false" autocomplete="off" value="' + (data ? h(data.name) : '') + '"/></div>' +
           '<div class="form-row"><label>Group</label>' +
           '<input id="df-group" class="field" value="Adhoc" readonly/></div>' +
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

// ── Save handlers ──────────────────────────────────────────────────────────────

function saveScalar() {
  const name  = (document.getElementById('df-name')  || {}).value || '';
  const type  = (document.getElementById('df-type')  || {}).value || 'STRING';
  const value = (document.getElementById('df-value') || {}).value || '';
  const group = 'Adhoc';

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
  const group = 'Adhoc';
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
