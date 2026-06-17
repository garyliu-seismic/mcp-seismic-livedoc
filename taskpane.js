'use strict';

// ─── State ────────────────────────────────────────────────────────────────────
/** @type {Array<{name:string, type:string, defaultValue:string, group:string}>} */
let variables = [];

/** @type {Array<{name:string, group:string, columns:string[], rows:string[][]}>} */
let tableVariables = [];

const XML_NS = 'http://schemas.livedoc.seismic.com/poc-variables/v1';

// ─── Type badge labels ────────────────────────────────────────────────────────
const TYPE_LABELS = { STRING: 'ABC', NUMBER: '123', DATE: 'DT', BOOLEAN: 'T/F' };

// ─── Office.js init ───────────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('app').style.display     = 'flex';
    setupTabs();
    loadFromCustomXml();
  } else {
    document.getElementById('loading').innerHTML =
      '<p style="color:#c00;padding:20px">This add-in requires PowerPoint.</p>';
  }
});

// ─── Tab wiring ───────────────────────────────────────────────────────────────
function setupTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    });
  });
}

// ─── Custom XML — Load ────────────────────────────────────────────────────────
// Must use PowerPoint.run + ctx.presentation.customXmlParts (PowerPointApi 1.7).
// Office.context.document.customXmlParts is Common API and is undefined in PPT.
async function loadFromCustomXml() {
  showStatus('Loading from document…', 'info');
  try {
    await PowerPoint.run(async (ctx) => {
      const parts = ctx.presentation.customXmlParts;
      parts.load('items');
      await ctx.sync();

      console.log('[POC] Total customXmlParts:', parts.items.length);
      if (parts.items.length === 0) {
        showStatus('No saved variables in this document yet.', 'info');
        renderAll();
        return;
      }

      // getXml() returns a ClientResult<string> — must call before sync
      const xmlResults = parts.items.map(p => p.getXml());
      await ctx.sync();

      const xmlStrings = xmlResults.map(r => r.value || '');
      console.log('[POC] XML values:', xmlStrings.map(s => s.slice(0, 120)));

      const matchXml = xmlStrings.find(s => s.includes(XML_NS));
      if (!matchXml) {
        showStatus('No saved variables in this document yet.', 'info');
        renderAll();
        return;
      }

      try {
        parseXml(matchXml);
        showStatus(`Loaded ${variables.length} variable(s) and ${tableVariables.length} table variable(s).`, 'success');
      } catch (e) {
        console.error('[POC] Parse error:', e.message, '\nXML:', matchXml);
        showStatus('Could not parse saved data: ' + e.message, 'error');
      }
      renderAll();
    });
  } catch (e) {
    showStatus('Load error: ' + e.message, 'error');
    renderAll();
  }
}

function refreshFromDoc() {
  variables = [];
  tableVariables = [];
  loadFromCustomXml();
}

// ─── Custom XML — Save ────────────────────────────────────────────────────────
async function saveToCustomXml(successMsg) {
  const xml = buildXml();
  try {
    await PowerPoint.run(async (ctx) => {
      const parts = ctx.presentation.customXmlParts;
      parts.load('items');
      await ctx.sync();

      const xmlResults = parts.items.map(p => p.getXml());
      await ctx.sync();

      const matchIdx = xmlResults.findIndex(r => (r.value || '').includes(XML_NS));
      if (matchIdx >= 0) {
        parts.items[matchIdx].setXml(xml);
      } else {
        parts.add(xml);
      }
      await ctx.sync();
      showStatus(successMsg || 'Saved.', 'success');
    });
  } catch (e) {
    showStatus('Save failed: ' + e.message, 'error');
  }
}

// ─── XML serialization ────────────────────────────────────────────────────────
function buildXml() {
  const varEls = variables.map(v =>
    `<variable name="${x(v.name)}" type="${x(v.type||'STRING')}" defaultValue="${x(v.defaultValue)}" group="${x(v.group||'')}"/>`
  ).join('\n    ');

  const tvEls = tableVariables.map(tv => {
    const cols = tv.columns.map(c => `<col>${x(c)}</col>`).join('');
    const rows = tv.rows.map(row =>
      `<row>${row.map(cell => `<cell>${x(cell)}</cell>`).join('')}</row>`
    ).join('\n        ');
    return `<tableVariable name="${x(tv.name)}" group="${x(tv.group||'')}">` +
           `<columns>${cols}</columns>` +
           `<rows>\n        ${rows}\n      </rows>` +
           `</tableVariable>`;
  }).join('\n    ');

  return `<?xml version="1.0" encoding="UTF-8"?>\n` +
         `<livedocVariables xmlns="${XML_NS}">\n` +
         `  <variables>\n    ${varEls}\n  </variables>\n` +
         `  <tableVariables>\n    ${tvEls}\n  </tableVariables>\n` +
         `</livedocVariables>`;
}

function parseXml(xmlStr) {
  if (!xmlStr || typeof xmlStr !== 'string' || !xmlStr.trim()) throw new Error('Empty XML');
  const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  if (doc.querySelector('parsererror')) throw new Error('Malformed XML');

  // Use getElementsByTagName (namespace-agnostic) — querySelectorAll fails on namespaced XML
  variables = [];
  Array.from(doc.getElementsByTagName('variable')).forEach(el => {
    variables.push({
      name:         el.getAttribute('name')         || '',
      type:         el.getAttribute('type')         || 'STRING',
      defaultValue: el.getAttribute('defaultValue') || '',
      group:        el.getAttribute('group')        || '',
    });
  });

  tableVariables = [];
  Array.from(doc.getElementsByTagName('tableVariable')).forEach(el => {
    tableVariables.push({
      name:    el.getAttribute('name')  || '',
      group:   el.getAttribute('group') || '',
      columns: Array.from(el.getElementsByTagName('col')).map(c => c.textContent),
      rows:    Array.from(el.getElementsByTagName('row')).map(row =>
                 Array.from(row.getElementsByTagName('cell')).map(c => c.textContent)),
    });
  });
}

// ─── Insert {{token}} at cursor ───────────────────────────────────────────────
function insertVariable(name) {
  const token = `{{${name}}}`;
  Office.context.document.setSelectedDataAsync(
    token,
    { coercionType: Office.CoercionType.Text },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showStatus(`Inserted: ${token}`, 'success');
      } else {
        showStatus(`Insert failed: ${result.error.message} — click inside a text box first.`, 'error');
      }
    }
  );
}

// ─── Insert table as PPT table shape ─────────────────────────────────────────
async function insertTableVariable(name) {
  const tv = tableVariables.find(t => t.name === name);
  if (!tv) { showStatus('Table variable not found.', 'error'); return; }

  const rowCount  = tv.rows.length + 1;
  const colCount  = tv.columns.length;
  const tblWidth  = Math.min(680, Math.max(280, colCount * 130));
  const tblHeight = rowCount * 36;

  try {
    await PowerPoint.run(async (context) => {
      let slide;
      try {
        const sel = context.presentation.getSelectedSlides();
        sel.load('items');
        await context.sync();
        slide = sel.items[0];
      } catch (_e) {
        slide = context.presentation.slides.getItemAt(0);
        showStatus('Inserting on slide 1 (getSelectedSlides unsupported)', 'info');
      }

      const shape = slide.shapes.addTable(rowCount, colCount, {
        left: 50, top: 120, width: tblWidth, height: tblHeight,
      });

      // Must sync before accessing shape.table rows
      await context.sync();

      const tbl = shape.table;
      for (let c = 0; c < colCount; c++) {
        tbl.rows.getItemAt(0).cells.getItemAt(c).text = tv.columns[c] || '';
      }
      for (let r = 0; r < tv.rows.length; r++) {
        for (let c = 0; c < colCount; c++) {
          tbl.rows.getItemAt(r + 1).cells.getItemAt(c).text = tv.rows[r][c] || '';
        }
      }

      await context.sync();
      showStatus(`Table "${name}" inserted (${rowCount} rows × ${colCount} cols)`, 'success');
    });
  } catch (err) {
    showStatus(`Table insert failed: ${err.message}`, 'error');
  }
}

// ─── Variable CRUD ────────────────────────────────────────────────────────────
function addVariable() {
  const name  = document.getElementById('varName').value.trim();
  const type  = document.getElementById('varType').value;
  const value = document.getElementById('varValue').value.trim();
  const group = document.getElementById('varGroup').value.trim();

  if (!name) { showStatus('Variable name is required.', 'error'); return; }
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    showStatus('Name must start with a letter/underscore and contain only A-Z, 0-9, _.', 'error');
    return;
  }
  if (variables.find(v => v.name === name)) {
    showStatus(`Variable "${name}" already exists.`, 'error'); return;
  }

  variables.push({ name, type, defaultValue: value, group });
  document.getElementById('varName').value  = '';
  document.getElementById('varValue').value = '';
  document.getElementById('varGroup').value = '';

  saveToCustomXml(`Variable "${name}" added.`);
  renderVariableTree();
}

function deleteVariable(name) {
  variables = variables.filter(v => v.name !== name);
  saveToCustomXml(`Variable "${name}" deleted.`);
  renderVariableTree();
}

// ─── Table variable CRUD ──────────────────────────────────────────────────────
function addTableVariable() {
  const name  = document.getElementById('tvName').value.trim();
  const group = document.getElementById('tvGroup').value.trim();
  const csv   = document.getElementById('tvCsv').value.trim();

  if (!name) { showStatus('Table variable name is required.', 'error'); return; }
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    showStatus('Name must start with a letter/underscore and contain only A-Z, 0-9, _.', 'error');
    return;
  }
  if (tableVariables.find(t => t.name === name)) {
    showStatus(`Table variable "${name}" already exists.`, 'error'); return;
  }
  if (!csv) { showStatus('CSV data is required.', 'error'); return; }

  const lines   = csv.split('\n').map(l => l.trim()).filter(Boolean);
  const columns = lines[0].split(',').map(c => c.trim());
  const rows    = lines.slice(1).map(line => {
    const cells = line.split(',').map(c => c.trim());
    while (cells.length < columns.length) cells.push('');
    return cells.slice(0, columns.length);
  });

  tableVariables.push({ name, group, columns, rows });
  document.getElementById('tvName').value  = '';
  document.getElementById('tvGroup').value = '';
  document.getElementById('tvCsv').value   = '';

  saveToCustomXml(`Table variable "${name}" added.`);
  renderTableVariableTree();
}

function deleteTableVariable(name) {
  tableVariables = tableVariables.filter(t => t.name !== name);
  saveToCustomXml(`Table variable "${name}" deleted.`);
  renderTableVariableTree();
}

// ─── Rendering ────────────────────────────────────────────────────────────────
function renderAll() {
  renderVariableTree();
  renderTableVariableTree();
}

function renderVariableTree() {
  const container = document.getElementById('varTree');
  document.getElementById('varCount').textContent = variables.length;

  if (variables.length === 0) {
    container.innerHTML = `
      <div class="tree-empty">
        <div class="tree-empty-icon">📋</div>
        No variables yet — add one above.
      </div>`;
    return;
  }

  const grouped = groupBy(variables, v => v.group || '');
  container.innerHTML = buildGroupedTree(grouped, renderVarRow);
  wireGroupToggles(container);
}

function renderVarRow(v) {
  const type  = v.type || 'STRING';
  const label = TYPE_LABELS[type] || type.slice(0, 3);
  const val   = v.defaultValue ? h(v.defaultValue) : '<em style="color:#bbb">—</em>';
  return `
    <div class="tree-var-row">
      <span class="type-badge badge-${h(type)}" title="${h(type)}">${label}</span>
      <div class="var-info">
        <div class="var-name">${h(v.name)}</div>
        <div class="var-value">${val}</div>
      </div>
      <div class="tree-actions">
        <button class="act-btn act-insert" onclick="insertVariable('${j(v.name)}')" title="Insert at cursor">→</button>
        <button class="act-btn act-delete" onclick="deleteVariable('${j(v.name)}')" title="Delete">✕</button>
      </div>
    </div>`;
}

function renderTableVariableTree() {
  const container = document.getElementById('tvTree');
  document.getElementById('tvCount').textContent = tableVariables.length;

  if (tableVariables.length === 0) {
    container.innerHTML = `
      <div class="tree-empty">
        <div class="tree-empty-icon">🗃</div>
        No table variables yet — add one above.
      </div>`;
    return;
  }

  const grouped = groupBy(tableVariables, t => t.group || '');
  container.innerHTML = buildGroupedTree(grouped, renderTvRow);
  wireGroupToggles(container);
}

function renderTvRow(tv) {
  const preview = tv.columns.length > 0 ? `${tv.columns.length} cols × ${tv.rows.length} rows` : '';
  return `
    <div class="tree-var-row">
      <span class="type-badge badge-TABLE" title="TABLE">TBL</span>
      <div class="var-info">
        <div class="var-name">${h(tv.name)}</div>
        <div class="var-value">${preview}</div>
      </div>
      <div class="tree-actions">
        <button class="act-btn act-insert" onclick="insertTableVariable('${j(tv.name)}')" title="Insert table">→</button>
        <button class="act-btn act-delete" onclick="deleteTableVariable('${j(tv.name)}')" title="Delete">✕</button>
      </div>
    </div>`;
}

// ─── Tree helpers ─────────────────────────────────────────────────────────────
function groupBy(arr, keyFn) {
  const map = new Map();
  arr.forEach(item => {
    const k = keyFn(item);
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(item);
  });
  return map;
}

function buildGroupedTree(grouped, renderItem) {
  let html = '';
  const ungrouped   = grouped.get('') || [];
  const namedGroups = [...grouped.entries()].filter(([k]) => k !== '');

  namedGroups.forEach(([groupName, items]) => {
    html += `
      <div class="tree-group">
        <div class="tree-group-header">
          <span class="folder-toggle">▾</span>
          <span class="folder-icon">📁</span>
          <span class="group-name">${h(groupName)}</span>
          <span class="group-count">${items.length}</span>
        </div>
        <div class="tree-group-body">
          ${items.map(renderItem).join('')}
        </div>
      </div>`;
  });

  ungrouped.forEach(item => { html += renderItem(item); });
  return html;
}

function wireGroupToggles(container) {
  container.querySelectorAll('.tree-group-header').forEach(header => {
    header.addEventListener('click', () => {
      const body   = header.nextElementSibling;
      const toggle = header.querySelector('.folder-toggle');
      const collapsed = body.classList.toggle('collapsed');
      toggle.textContent = collapsed ? '▸' : '▾';
    });
  });
}

// ─── Status bar ───────────────────────────────────────────────────────────────
let _statusTimer = null;

function showStatus(msg, type = 'info') {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className   = `status-bar ${type}`;
  el.style.display = 'block';
  clearTimeout(_statusTimer);
  if (type !== 'error') {
    _statusTimer = setTimeout(() => { el.style.display = 'none'; }, 4000);
  }
}

// ─── Escape helpers ───────────────────────────────────────────────────────────
function x(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
                  .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}
function h(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function j(s) {
  return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'");
}
