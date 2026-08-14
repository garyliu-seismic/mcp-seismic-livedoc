import { randomUUID } from "crypto";
import { dbg } from "../utils/debug.js";
import { BROWSER_CLIENT_ID, BROWSER_SCOPES, DEFAULT_AUTH_URI } from "../config.js";

function cookieStr(headers: Headers): string {
  const cookies = (headers as unknown as { getSetCookie?(): string[] }).getSetCookie?.() ?? [];
  return cookies.map(c => c.split(";")[0]).join("; ");
}

export async function browserLogin(tenant: string, username: string, password: string, authUri = DEFAULT_AUTH_URI): Promise<string> {
  const authBase    = `${authUri}/tenants/${encodeURIComponent(tenant)}`;
  const redirectUri = `https://${tenant}.seismic.com/app`;
  const state    = randomUUID().replace(/-/g, "");
  const nonce    = randomUUID().replace(/-/g, "");
  const appState = randomUUID().replace(/-/g, "");

  const params = new URLSearchParams({
    client_id: BROWSER_CLIENT_ID, response_type: "id_token token", scope: BROWSER_SCOPES,
    state, redirect_uri: `${redirectUri}?state=${appState}`,
    response_mode: "form_post", nonce, themeMode: "light",
  });

  dbg(`browserLogin: step1 tenant=${tenant} authBase=${authBase}`);
  // Step 1: GET /connect/authorize with redirect:manual — the session cookie is on this first
  // 302 response itself. We do NOT follow the redirect (it goes to the tenant login page which
  // may be unreachable). The cookie from this response is all we need for step 2.
  const step1 = await fetch(`${authBase}/connect/authorize?${params}`, {
    redirect: "manual",
    headers: { "User-Agent": "Mozilla/5.0" },
  }).catch(e => { throw new Error(`Step 1 (authorize) network error: ${e}`); });
  const cookies1 = cookieStr(step1.headers);
  dbg(`browserLogin: step1 status=${step1.status} cookies=${cookies1 ? cookies1.slice(0, 80) : "(none)"}`);
  if (!cookies1) throw new Error(`Step 1 (authorize) returned no cookies (status ${step1.status}). Auth server may be unreachable or the client_id is not registered for this tenant.`);

  // Step 2: POST credentials
  const loginRes = await fetch(`${authBase}/api/v1/account/login`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "User-Agent": "Mozilla/5.0", Cookie: cookies1 },
    body: JSON.stringify({ Username: username, Password: password, RememberMe: false, ClientId: BROWSER_CLIENT_ID, ClientVersion: "", DisableSingleSignOn: false }),
  }).catch(e => { throw new Error(`Step 2 (login POST) fetch error: ${e}`); });
  const loginText = await loginRes.text().catch(e => { throw new Error(`Step 2 (login POST) body read error: ${e}`); });
  let loginData: { isSuccess: boolean };
  try { loginData = JSON.parse(loginText) as { isSuccess: boolean }; }
  catch { throw new Error(`Step 2 (login POST) non-JSON response (HTTP ${loginRes.status}): ${loginText.slice(0, 300)}`); }
  dbg(`browserLogin: step2 status=${loginRes.status} isSuccess=${loginData.isSuccess}`);
  if (!loginData.isSuccess) throw new Error("Seismic login failed - check username/password");
  const cookies2 = [cookies1, cookieStr(loginRes.headers)].filter(Boolean).join("; ");

  // Step 3: callback — token is in the HTML form response
  const cbRes = await fetch(`${authBase}/connect/authorize/callback?${params}`, {
    redirect: "manual",
    headers: { "User-Agent": "Mozilla/5.0", Cookie: cookies2, Accept: "text/html" },
  }).catch(e => { throw new Error(`Step 3 (callback) fetch error: ${e}`); });
  const html = await cbRes.text();
  const m = html.match(/name=['"]access_token['"]\s+value=['"]([^'"]+)['"]/)
    ?? html.match(/value=['"]([^'"]+)['"]\s+name=['"]access_token['"]/);
  dbg(`browserLogin: step3 status=${cbRes.status} tokenFound=${!!m} bodySnippet=${html.slice(0, 100)}`);
  if (!m) throw new Error(`Could not extract access_token. Callback status: ${cbRes.status}, body snippet: ${html.slice(0, 200)}`);
  return m[1];
}
