'use strict';

// Chat mode for local-model orchestration.
// It sends user instructions to /api/chat-intent and executes returned actions
// using existing variable CRUD and dynamic binding functions.

function initChatMode() {
  var sendBtn = document.getElementById('chatSendBtn');
  var input = document.getElementById('chatInput');
  var clearBtn = document.getElementById('chatClearBtn');

  if (!sendBtn || !input) return;

  sendBtn.addEventListener('click', runChatCommand);
  input.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      runChatCommand();
    }
  });

  if (clearBtn) {
    clearBtn.addEventListener('click', function () {
      var log = document.getElementById('chatLog');
      if (log) log.innerHTML = '';
    });
  }

  appendChatLine('assistant', 'Chat mode is ready. Try: create a table variable RevenueByRegion and bind selected chart.');
}

async function runChatCommand() {
  var input = document.getElementById('chatInput');
  var sendBtn = document.getElementById('chatSendBtn');
  var text = input ? input.value.trim() : '';
  if (!text) return;

  appendChatLine('user', text);
  input.value = '';
  setChatBusy(true);
  if (sendBtn) sendBtn.disabled = true;

  try {
    var response = await fetch('/api/chat-intent', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: text,
        variables: buildVariableContextSnapshot(),
      }),
    });

    var payload = await response.json();
    if (!response.ok) {
      throw new Error(payload && payload.detail ? payload.detail : 'chat-intent request failed');
    }

    var actions = Array.isArray(payload.actions) ? payload.actions : [];
    var result = await executePlannedActions(actions);

    var summary = payload.assistantMessage || 'Plan created.';
    if (actions.length) {
      summary += ' Executed ' + actions.length + ' action(s): ' + result.ok + ' succeeded';
      if (result.failed) summary += ', ' + result.failed + ' failed';
      summary += '.';
    }
    appendChatLine('assistant', summary);

    if (result.messages.length) {
      result.messages.forEach(function (m) {
        appendChatLine(m.level === 'error' ? 'assistant-error' : 'assistant', m.text);
      });
    }
  } catch (e) {
    appendChatLine('assistant-error', 'Chat mode failed: ' + e.message);
  } finally {
    setChatBusy(false);
    if (sendBtn) sendBtn.disabled = false;
  }
}

function buildVariableContextSnapshot() {
  return (variables || []).map(function (v) {
    return {
      name: v.name,
      kind: v.kind,
      type: v.type,
      group: v.group,
      columns: v.columns || [],
      rowCount: (v.rows || []).length,
    };
  });
}

async function executePlannedActions(actions) {
  var messages = [];
  var ok = 0;
  var failed = 0;
  var needsSave = false;

  for (var i = 0; i < actions.length; i++) {
    var action = actions[i] || {};
    try {
      var changed = await executeAction(action, messages);
      if (changed) needsSave = true;
      ok++;
    } catch (e) {
      failed++;
      messages.push({ level: 'error', text: 'Action ' + (action.type || '(unknown)') + ' failed: ' + e.message });
    }
  }

  if (needsSave) {
    await saveToCustomXml('Chat mode applied updates.');
    renderTree();
  }

  return { ok: ok, failed: failed, messages: messages };
}

async function executeAction(action, messages) {
  var type = String(action.type || '').trim();

  if (type === 'create_scalar_variable') {
    var scalar = createScalarVariableFromAction(action);
    messages.push({ level: 'info', text: 'Variable created: ' + scalar.name });
    return true;
  }

  if (type === 'create_table_variable') {
    var table = createTableVariableFromAction(action);
    messages.push({ level: 'info', text: 'Table variable created: ' + table.name + ' (' + (table.rows || []).length + ' rows).' });
    return true;
  }

  if (type === 'create_computed_variable') {
    var computed = createComputedVariableFromAction(action);
    messages.push({ level: 'info', text: 'Computed variable created: ' + computed.name });
    return true;
  }

  if (type === 'insert_variable') {
    var v = findVariableByName(action.name);
    if (!v) throw new Error('Variable not found: ' + action.name);
    insertVariable(v.id);
    messages.push({ level: 'info', text: 'Inserted variable token/table: ' + v.name });
    return false;
  }

  if (type === 'insert_table') {
    var tv = findVariableByName(action.name);
    if (!tv || tv.kind !== 'table') throw new Error('Table variable not found: ' + action.name);
    await insertTableShape(tv);
    messages.push({ level: 'info', text: 'Inserted table shape from variable: ' + tv.name });
    return false;
  }

  if (type === 'configure_dynamic_table') {
    await setShapeTagOnCurrentSlide(String(action.shapeName || ''), 'LIVEDOC_DYN_TABLE', {
      variableName: String(action.tableVar || ''),
      fromRow: Math.max(1, parseInt(action.fromRow || '1', 10) || 1),
      toRow: Math.max(1, parseInt(action.toRow || '1', 10) || 1),
    });
    messages.push({ level: 'info', text: 'Dynamic table binding saved for shape: ' + action.shapeName });
    return false;
  }

  if (type === 'configure_dynamic_chart') {
    await setShapeTagOnCurrentSlide(String(action.shapeName || ''), 'LIVEDOC_DYN_CHART', {
      titleVar: String(action.titleVar || ''),
      tableVar: String(action.tableVar || ''),
      labelCol: String(action.labelCol || ''),
      valueCol: String(action.valueCol || ''),
    });
    messages.push({ level: 'info', text: 'Dynamic chart binding saved for shape: ' + action.shapeName });
    return false;
  }

  if (type === 'configure_dynamic_image') {
    await setShapeTagOnCurrentSlide(String(action.shapeName || ''), 'LIVEDOC_DYN_IMAGE', {
      fitMode: String(action.fitMode || 'fitInside'),
      sourceUrl: String(action.sourceUrl || ''),
    });
    messages.push({ level: 'info', text: 'Dynamic image binding saved for shape: ' + action.shapeName });
    return false;
  }

  if (type === 'run_preview') {
    await previewDoc();
    messages.push({ level: 'info', text: 'Preview command executed.' });
    return false;
  }

  throw new Error('Unsupported action type: ' + type);
}

function findVariableByName(name) {
  var key = String(name || '').toLowerCase();
  return (variables || []).find(function (v) { return String(v.name || '').toLowerCase() === key; }) || null;
}

function normalizeVarName(name) {
  var raw = String(name || '').trim().replace(/\s+/g, '_');
  var cleaned = raw.replace(/[^A-Za-z0-9_]/g, '_');
  if (!/^[A-Za-z_]/.test(cleaned)) cleaned = '_' + cleaned;
  return cleaned || ('Var_' + uid().slice(0, 6));
}

function normalizeScalarType(dataType) {
  var t = String(dataType || 'STRING').toUpperCase();
  if (t !== 'STRING' && t !== 'NUMBER' && t !== 'DATE' && t !== 'BOOLEAN') return 'STRING';
  return t;
}

function createScalarVariableFromAction(action) {
  var name = normalizeVarName(action.name || 'NewVariable');
  var existing = findVariableByName(name);
  var entry = {
    id: existing ? existing.id : uid(),
    kind: 'scalar',
    name: name,
    type: normalizeScalarType(action.dataType),
    defaultValue: action.defaultValue == null ? '' : String(action.defaultValue),
    group: String(action.group || 'Adhoc'),
  };

  if (existing) {
    var idx = variables.findIndex(function (v) { return v.id === existing.id; });
    if (idx >= 0) variables[idx] = entry;
  } else {
    variables.push(entry);
  }

  return entry;
}

function createComputedVariableFromAction(action) {
  var name = normalizeVarName(action.name || 'ComputedVariable');
  var existing = findVariableByName(name);
  var entry = {
    id: existing ? existing.id : uid(),
    kind: 'computed',
    name: name,
    formula: String(action.formula || ''),
    group: String(action.group || 'Computed'),
  };

  if (existing) {
    var idx = variables.findIndex(function (v) { return v.id === existing.id; });
    if (idx >= 0) variables[idx] = entry;
  } else {
    variables.push(entry);
  }

  return entry;
}

function normalizeTableColumns(rawColumns) {
  if (Array.isArray(rawColumns)) {
    return rawColumns.map(function (c) { return String(c || '').trim(); }).filter(Boolean);
  }
  if (typeof rawColumns === 'string') {
    return rawColumns.split(',').map(function (c) { return c.trim(); }).filter(Boolean);
  }
  return [];
}

function normalizeTableRows(rawRows, columns) {
  if (!Array.isArray(rawRows)) return [];

  return rawRows.map(function (row) {
    // Accept string row: "APAC,120" or "APAC 120"
    if (typeof row === 'string') {
      var split = row.includes(',')
        ? row.split(',')
        : row.trim().split(/\s+/);
      var outFromString = [];
      for (var si = 0; si < columns.length; si++) {
        outFromString.push(split[si] == null ? '' : String(split[si]).trim());
      }
      return outFromString;
    }

    // Accept array rows: ["APAC", 120]
    if (Array.isArray(row)) {
      // Also accept wrapped object row: [{Region:"APAC",Amount:120}]
      if (row.length === 1 && row[0] && typeof row[0] === 'object' && !Array.isArray(row[0])) {
        var wrappedObj = row[0];
        var outFromWrapped = [];
        for (var wi = 0; wi < columns.length; wi++) {
          var wk = columns[wi];
          outFromWrapped.push(wrappedObj[wk] == null ? '' : String(wrappedObj[wk]));
        }
        return outFromWrapped;
      }

      var outFromArray = [];
      for (var i = 0; i < columns.length; i++) {
        outFromArray.push(row[i] == null ? '' : String(row[i]));
      }
      return outFromArray;
    }

    // Accept object rows: {Region: "APAC", Amount: 120}
    if (row && typeof row === 'object') {
      var outFromObject = [];
      for (var j = 0; j < columns.length; j++) {
        var key = columns[j];
        outFromObject.push(row[key] == null ? '' : String(row[key]));
      }
      return outFromObject;
    }

    // Fallback single cell value
    return [String(row)];
  }).map(function (cells) {
    // Ensure each row length matches columns length.
    var out = cells.slice(0, columns.length);
    while (out.length < columns.length) out.push('');
    return out;
  });
}

function createTableVariableFromAction(action) {
  var name = normalizeVarName(action.name || 'TableVariable');
  var columns = normalizeTableColumns(action.columns);

  // If model omitted columns but returned object rows, derive columns from first object row.
  if (!columns.length && Array.isArray(action.rows) && action.rows.length) {
    var firstRow = action.rows[0];
    if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
      columns = Object.keys(firstRow);
    }
  }

  if (!columns.length) columns = ['Column1', 'Column2'];

  var rows = normalizeTableRows(action.rows, columns);

  var existing = findVariableByName(name);
  var existingTypes = existing && existing.columnTypes ? existing.columnTypes : [];
  var columnTypes = columns.map(function (_, i) {
    return existingTypes[i] || 'STRING';
  });

  var entry = {
    id: existing ? existing.id : uid(),
    kind: 'table',
    name: name,
    group: String(action.group || 'Adhoc'),
    columns: columns,
    columnTypes: columnTypes,
    rows: rows,
  };

  if (existing) {
    var idx = variables.findIndex(function (v) { return v.id === existing.id; });
    if (idx >= 0) variables[idx] = entry;
  } else {
    variables.push(entry);
  }

  return entry;
}

async function setShapeTagOnCurrentSlide(shapeName, tagKey, tagValueObject) {
  if (!shapeName) throw new Error('shapeName is required');
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
    if (!shape) throw new Error('Shape not found on current slide: ' + shapeName);

    shape.tags.add(tagKey, JSON.stringify(tagValueObject || {}));
    await context.sync();
  });
}

function appendChatLine(role, text) {
  var log = document.getElementById('chatLog');
  if (!log) return;
  var line = document.createElement('div');
  line.className = 'chat-line ' + role;
  line.textContent = text;
  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
}

function setChatBusy(isBusy) {
  var indicator = document.getElementById('chatBusy');
  if (!indicator) return;
  indicator.style.display = isBusy ? 'inline-block' : 'none';
}
