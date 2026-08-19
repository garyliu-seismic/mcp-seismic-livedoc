'use strict';

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

// ── Insert token / variable into the open document ────────────────────────────

// Insert a raw {{token}} string at the current cursor position.
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

// Insert either a {{token}} (scalar/computed) or a live table shape (table/system).
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

// Insert a table shape on the current slide populated with the variable's data.
async function insertTableShape(v) {
  const columns  = v.columns  || [];
  const colCount = columns.length;
  if (colCount === 0) { showStatus('Table has no columns defined.', 'error'); return; }

  // Default insertion is preview-ready dynamic table:
  // Row 1 = header, Row 2 = template row with child variable tokens.
  const rowCount  = 2;
  const tblWidth  = Math.min(680, Math.max(280, colCount * 130));
  const tblHeight = Math.max(40, rowCount * 36);

  var shapeName = 'DynTable_' + (v.name || 'Table') + '_' + uid().slice(0, 6);

  // Check whether PowerPointApi 1.3 (shape.table) is available
  var hasTableApi = Office.context.requirements.isSetSupported('PowerPointApi', '1.3');

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

      var shape = slide.shapes.addTable(rowCount, colCount, {
        left: 50, top: 120, width: tblWidth, height: tblHeight,
      });
      await context.sync();

      shape.name = shapeName;

      // autoTokenize: preview engine fills header+data by column index (no tokens needed in cells)
      shape.tags.add('LIVEDOC_DYN_TABLE', JSON.stringify({
        variableName: v.name,
        fromRow: 2,
        toRow: 2,
        autoTokenize: true,
        conditionalRows: [],
        mergeRules: [],
      }));
      await context.sync();

      showStatus('Table "' + v.name + '" inserted (' + colCount + ' cols). Click Preview to expand with data.', 'success');
    });
  } catch (err) {
    showStatus('Table insert failed: ' + err.message, 'error');
  }
}
