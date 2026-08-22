import S from "../styles.js";

export function ScalarInput({ field, value, onChange, error }) {
  const t = (field.type || "STRING").toUpperCase();
  const required = !!field.validation?.required;
  // Prefer the template author's own BasicSetting.Label over the raw bound variable name.
  const displayName = field.label || field.name;
  const label = (
    <label style={S.fl} title={field.tooltip || undefined}>
      {displayName}{required && <span style={{ color: "#c00" }}> *</span>}
    </label>
  );
  const helpText = field.helpText && <div style={S.helpText}>{field.helpText}</div>;

  if (t === "BOOL" || t === "BOOLEAN") {
    return (
      <div style={S.fw}>
        <label style={S.bool} title={field.tooltip || undefined}>
          <input type="checkbox" checked={!!value} onChange={e => onChange(e.target.checked)} />
          <span style={S.fl}>{displayName}{required && <span style={{ color: "#c00" }}> *</span>}</span>
        </label>
        {helpText}
      </div>
    );
  }

  // Static domain-of-value list (fixed choices from the form definition) — render as a
  // dropdown instead of free text.
  if (field.options?.length > 0) {
    return (
      <div style={S.fw}>
        {label}
        {helpText}
        <select
          value={value ?? ""}
          onChange={e => onChange(e.target.value)}
          style={{ ...S.fi, ...(error ? S.fiErr : {}) }}
        >
          <option value="" disabled={required}>{required ? "Select…" : "(none)"}</option>
          {field.options.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
        </select>
        {error && <div style={S.fieldErr}>{error}</div>}
      </div>
    );
  }

  return (
    <div style={S.fw}>
      {label}
      {helpText}
      <input
        type={t === "DATE" ? "date" : (t === "INTEGER" || t === "FLOAT") ? "number" : "text"}
        step={t === "FLOAT" ? "any" : undefined}
        value={value ?? ""}
        onChange={e => onChange(e.target.value)}
        style={{ ...S.fi, ...(error ? S.fiErr : {}) }}
      />
      {error && <div style={S.fieldErr}>{error}</div>}
    </div>
  );
}
