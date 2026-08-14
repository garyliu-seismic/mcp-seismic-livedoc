import { useState, useEffect, useRef } from "react";
import { App } from "@modelcontextprotocol/ext-apps";

export function useAppConnection({ onFormLoad, onError }) {
  const [phase, setPhase] = useState("connecting");
  const [loginError, setLoginError] = useState(null);
  const [loginBusy, setLoginBusy] = useState(false);
  const appRef = useRef(null);
  const tokenRef = useRef(null);
  const phaseRef = useRef("connecting");
  phaseRef.current = phase;

  // Keep callback refs current so effects don't capture stale closures
  const onFormLoadRef = useRef(onFormLoad);
  onFormLoadRef.current = onFormLoad;
  const onErrorRef = useRef(onError);
  onErrorRef.current = onError;

  async function handlePanelLogin(tenant, username, password) {
    setLoginBusy(true);
    setLoginError(null);
    try {
      const res = await appRef.current.callServerTool({ name: "panel_login", arguments: { tenant, username, password } });
      const sc = res?.structuredContent;
      if (sc?.ok) {
        setPhase("connecting");
      } else {
        setLoginError(sc?.error ?? "Login failed. Check your credentials.");
      }
    } catch (e) {
      setLoginError(String(e));
    } finally {
      setLoginBusy(false);
    }
  }

  async function loadSchema(app, formToken) {
    const res = await app.callServerTool({ name: "get_form_schema", arguments: { token: formToken } });
    const sc = res?.structuredContent;
    if (!sc || sc.error) throw new Error(sc?.error ?? "No schema returned");
    return sc;
  }

  useEffect(() => {
    const app = new App({ name: "livedoc-form", version: "1.0.0" }, {});
    appRef.current = app;

    app.ontoolresult = async (event) => {
      if (event?.structuredContent?.action === "show_login") { setPhase("login"); return; }
      if (event?.structuredContent?.action === "ready") { return; }

      const formToken = event?.structuredContent?.formToken;
      if (!formToken) return;
      // Don't hijack an active form — only accept a new schema when idle.
      if (phaseRef.current !== "connecting") return;
      tokenRef.current = formToken;
      phaseRef.current = "loading";
      setPhase("loading");
      try {
        const sc = await loadSchema(app, formToken);
        onFormLoadRef.current(sc);
        app.sendSizeChanged({ width: 520, height: 900 });
        setPhase("ready");
      } catch (e) {
        onErrorRef.current?.(String(e));
        setPhase("error");
      }
    };

    app.connect()
      .then(() => {
        app.callServerTool({ name: "get_auth_status", arguments: {} }).then(res => {
          if (!res?.structuredContent?.isAuthenticated) setPhase("login");
        }).catch(() => {});
      })
      .catch(e => { onErrorRef.current?.(String(e)); setPhase("error"); });
  }, []);

  // Poll for new generation requests when panel is idle
  useEffect(() => {
    const interval = setInterval(async () => {
      if (phaseRef.current !== "connecting") return;
      const app = appRef.current;
      if (!app) return;
      try {
        const authRes = await app.callServerTool({ name: "get_auth_status", arguments: {} });
        if (!authRes?.structuredContent?.isAuthenticated) { setPhase("login"); return; }

        const res = await app.callServerTool({
          name: "get_latest_token",
          arguments: { currentToken: tokenRef.current ?? "" },
        });
        const sc = res?.structuredContent;
        if (!sc?.isNew || !sc?.formToken) return;
        if (phaseRef.current !== "connecting") return;

        tokenRef.current = sc.formToken;
        phaseRef.current = "loading";
        setPhase("loading");

        const newSc = await loadSchema(app, sc.formToken);
        onFormLoadRef.current(newSc);
        app.sendSizeChanged({ width: 520, height: 900 });
        setPhase("ready");
      } catch { /* ignore poll errors */ }
    }, 2500);
    return () => clearInterval(interval);
  }, []);

  return { phase, setPhase, loginError, loginBusy, appRef, tokenRef, handlePanelLogin };
}
