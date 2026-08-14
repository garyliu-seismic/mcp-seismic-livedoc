export function emptyRow(columns) {
  const row = {};
  columns.forEach(c => {
    const t = (c.colType || "TEXT").toUpperCase();
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
