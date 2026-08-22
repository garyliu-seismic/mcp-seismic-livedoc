export function emptyRow(columns) {
  const row = {};
  columns.forEach(c => {
    const t = (c.colType || "TEXT").toUpperCase();
    if (c.defaultValue !== undefined) { row[c.name] = coerce(c.defaultValue, t); return; }
    row[c.name] = (t === "BOOL" || t === "BOOLEAN") ? false : "";
  });
  return row;
}

export function coerce(value, type) {
  const t = (type || "STRING").toUpperCase();
  if (t === "INTEGER") return value === "" || value === undefined ? 0 : (parseInt(value) || 0);
  if (t === "FLOAT")   return value === "" || value === undefined ? 0 : (parseFloat(value) || 0);
  if (t === "BOOL" || t === "BOOLEAN") return !!value;
  return value ?? "";
}

export function coerceRow(row, columns) {
  const out = {};
  columns.forEach(c => {
    const t = (c.colType || "TEXT").toUpperCase();
    out[c.name] = (t === "INTEGER") ? (parseInt(row[c.name]) || 0)
      : (t === "FLOAT") ? (parseFloat(row[c.name]) || 0)
      : (t === "BOOL" || t === "BOOLEAN") ? !!row[c.name]
      : (row[c.name] ?? "");
  });
  return out;
}

export function tableValue(rows, columns) {
  return {
    columns: columns.map(c => c.name),
    rows: rows.map(r => columns.map(c => coerceRow(r, columns)[c.name])),
  };
}

// Validates a single scalar value against the field's `validation` metadata (sourced from
// the form definition — see get_form_definition / formDefinition.ts). Returns an error
// message string, or null if the value is valid.
export function validateField(value, field) {
  const v = field?.validation;
  if (!v) return null;
  const t = (field.type || "STRING").toUpperCase();
  const isEmpty = value === undefined || value === null || value === "";
  if (v.required && isEmpty) return "Required";
  if (isEmpty) return null;

  if (t === "INTEGER" || t === "FLOAT") {
    const n = Number(value);
    if (!Number.isNaN(n)) {
      if (v.min != null && n < v.min) return `Must be at least ${v.min}`;
      if (v.max != null && n > v.max) return `Must be at most ${v.max}`;
    }
    return null;
  }

  const s = String(value);
  if (v.minLength != null && s.length < v.minLength) return `Must be at least ${v.minLength} characters`;
  if (v.maxLength != null && s.length > v.maxLength) return `Must be at most ${v.maxLength} characters`;
  if (v.pattern) {
    try {
      if (!new RegExp(v.pattern).test(s)) return v.patternMessage || "Invalid format";
    } catch { /* malformed pattern from the server — skip rather than block submit */ }
  }
  return null;
}

// Checks a table's required columns are filled in every row. Returns an error message
// string (naming the first offending column), or null if all required cells are filled.
export function validateTableRequired(rows, columns) {
  const requiredCols = columns.filter(c => c.validation?.required);
  if (requiredCols.length === 0) return null;
  for (let ri = 0; ri < rows.length; ri++) {
    for (const c of requiredCols) {
      const v = rows[ri][c.name];
      if (v === undefined || v === null || v === "") {
        return `"${c.name}" is required (row ${ri + 1})`;
      }
    }
  }
  return null;
}
