'use strict';

// ── Shared utility helpers ─────────────────────────────────────────────────────

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

// XML-safe attribute escape (for building XML strings)
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

// XML value escape (for injecting replacement text into OOXML)
function xmlEscapeVal(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
