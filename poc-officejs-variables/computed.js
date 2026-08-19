'use strict';

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

    var fn     = new Function(allKeys, '"use strict"; return (' + expr + ');');
    var result = fn.apply(null, allVals);

    if (result === null || result === undefined) return '';
    if (typeof result === 'object') return JSON.stringify(result);
    return String(result);
  } catch (e) {
    return '#ERR: ' + String(e.message).slice(0, 60);
  }
}
