import S from "../styles.js";
import { emptyRow } from "../utils/formUtils.js";

export function TableInput({ table, rows, onChange }) {
  const addRow = () => onChange([...rows, emptyRow(table.columns)]);
  const removeRow = i => onChange(rows.filter((_, idx) => idx !== i));
  const setCell = (ri, col, val) => onChange(rows.map((r, i) => i === ri ? { ...r, [col]: val } : r));

  return (
    <div style={S.tblWrap}>
      <div style={{ ...S.fl, marginBottom: 4 }}>{table.name}</div>
      <table style={S.table}>
        <thead>
          <tr>
            {table.columns.map(c => <th key={c.name} style={S.th}>{c.name}</th>)}
            <th style={{ ...S.th, width: 24 }} />
          </tr>
        </thead>
        <tbody>
          {rows.map((row, ri) => (
            <tr key={ri}>
              {table.columns.map(c => {
                const t = (c.colType || "TEXT").toUpperCase();
                const bool = t === "BOOL" || t === "BOOLEAN";
                return (
                  <td key={c.name} style={S.td}>
                    {bool
                      ? <input type="checkbox" checked={!!row[c.name]} onChange={e => setCell(ri, c.name, e.target.checked)} />
                      : <input
                          type={(t === "INTEGER" || t === "FLOAT") ? "number" : t === "DATE" ? "date" : "text"}
                          step={t === "FLOAT" ? "any" : undefined}
                          value={row[c.name] ?? ""}
                          onChange={e => setCell(ri, c.name, e.target.value)}
                          style={S.ci}
                        />
                    }
                  </td>
                );
              })}
              <td style={S.td}><button onClick={() => removeRow(ri)} style={S.delBtn}>×</button></td>
            </tr>
          ))}
        </tbody>
      </table>
      <button onClick={addRow} style={S.addBtn}>+ Add row</button>
    </div>
  );
}
