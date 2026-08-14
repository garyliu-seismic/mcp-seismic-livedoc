import S from "../styles.js";

export function ScalarInput({ field, value, onChange }) {
  const t = (field.type || "STRING").toUpperCase();
  if (t === "BOOL" || t === "BOOLEAN") {
    return (
      <label style={S.bool}>
        <input type="checkbox" checked={!!value} onChange={e => onChange(e.target.checked)} />
        <span style={S.fl}>{field.name}</span>
      </label>
    );
  }
  return (
    <div style={S.fw}>
      <label style={S.fl}>{field.name}</label>
      <input
        type={t === "DATE" ? "date" : (t === "INTEGER" || t === "FLOAT") ? "number" : "text"}
        step={t === "FLOAT" ? "any" : undefined}
        value={value ?? ""}
        onChange={e => onChange(e.target.value)}
        style={S.fi}
      />
    </div>
  );
}
