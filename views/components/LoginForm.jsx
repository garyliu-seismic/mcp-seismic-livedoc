import { useState, useEffect } from "react";
import S from "../styles.js";

export function LoginForm({ app, onSubmit, busy, error }) {
  const [tenant,   setTenant]   = useState("");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");

  useEffect(() => {
    app.callServerTool({ name: "get_auth_config", arguments: {} }).then(res => {
      const sc = res?.structuredContent;
      if (sc?.tenant) setTenant(sc.tenant);
    }).catch(() => {});
  }, []);

  function handleSubmit(e) {
    e.preventDefault();
    if (!tenant || !username || !password || busy) return;
    onSubmit(tenant, username, password);
  }
  const canSubmit = tenant && username && password && !busy;
  return (
    <div style={{ ...S.page, display: "flex", flexDirection: "column", gap: 0 }}>
      <div style={S.title}>Sign in to Seismic</div>
      <form onSubmit={handleSubmit} style={{ display: "flex", flexDirection: "column", gap: 12, maxWidth: 340 }}>
        <div style={S.fw}>
          <label style={S.fl}>Tenant</label>
          <input
            style={S.fi}
            type="text"
            placeholder="e.g. sttqaf12"
            value={tenant}
            onChange={e => setTenant(e.target.value)}
            disabled={busy}
          />
        </div>
        <div style={S.fw}>
          <label style={S.fl}>Username</label>
          <input
            style={S.fi}
            type="text"
            autoComplete="username"
            value={username}
            onChange={e => setUsername(e.target.value)}
            disabled={busy}
          />
        </div>
        <div style={S.fw}>
          <label style={S.fl}>Password</label>
          <input
            style={S.fi}
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={e => setPassword(e.target.value)}
            disabled={busy}
          />
        </div>
        {error && <div style={S.errBox}>{error}</div>}
        <button
          type="submit"
          disabled={!canSubmit}
          style={{ ...S.sub, opacity: canSubmit ? 1 : 0.6, cursor: busy ? "wait" : "pointer" }}
        >
          {busy
            ? <><span style={{ ...S.spinner, borderColor: "#fff", borderTopColor: "transparent" }} />Signing in…</>
            : "Sign in"
          }
        </button>
      </form>
    </div>
  );
}
