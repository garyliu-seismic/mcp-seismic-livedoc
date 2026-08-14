'use strict';

// ── Shared PPTX/OOXML traversal helpers ───────────────────────────────────────
// Used by dynamic-image.js and dynamic-chart.js.

// Find the first direct child element with a given localName (namespace-agnostic).
function findChildByLocalName(parent, localName) {
  for (var i = 0; i < parent.childNodes.length; i++) {
    var n = parent.childNodes[i];
    if (n.nodeType === 1 && n.localName === localName) return n;
  }
  return null;
}

// Scan a rels XML string for the Target attribute of a given relationship Id.
function findRelTargetById(relsXml, rId) {
  var escaped = rId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var idx = relsXml.search(new RegExp('Id="' + escaped + '"'));
  if (idx < 0) return null;
  var snippet = relsXml.slice(idx);
  var m = snippet.match(/Target="([^"]*)"/);
  return m ? m[1] : null;
}

// Resolve a relative path from a base directory.
// e.g. resolveRelativePath('ppt/slides/', '../charts/chart1.xml') → 'ppt/charts/chart1.xml'
function resolveRelativePath(baseDir, relTarget) {
  var parts = (baseDir + relTarget).split('/');
  var out = [];
  parts.forEach(function (p) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  });
  return out.join('/');
}

// Parse a self-contained XML string and import its root element into targetDoc.
// Used to create <c:strLit>/<c:numLit>/<c:rich> elements without namespace-prefix issues.
function importXmlFragment(targetDoc, xmlStr) {
  var fragDoc = new DOMParser().parseFromString(xmlStr, 'application/xml');
  if (fragDoc.getElementsByTagName('parseerror').length) return null;
  return targetDoc.importNode(fragDoc.documentElement, true);
}
