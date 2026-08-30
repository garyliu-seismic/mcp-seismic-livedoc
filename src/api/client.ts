import { BASE_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD } from "../config.js";
import { authHeaders } from "../auth/headers.js";
import { loadSavedToken } from "../auth/token-store.js";
import { getToken, setToken, isTokenManual, getCachedUsername, getCachedPassword } from "../auth/state.js";
import { autoLogin } from "../auth/auto-login.js";
import { dbg } from "../utils/debug.js";

const REQUEST_TIMEOUT_MS = 30_000;

export async function apiFetch(
  path: string,
  options: RequestInit = {},
  _retry = true,
  base: string = BASE_URL
): Promise<{ status: number; body: unknown }> {
  const url = `${base}${path}`;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  let res: Response;
  try {
    res = await fetch(url, {
      ...options,
      headers: { ...authHeaders(), ...(options.headers as Record<string, string> ?? {}) },
      signal: controller.signal,
    });
  } catch (err: unknown) {
    if (err instanceof Error && err.name === "AbortError") {
      return { status: 0, body: `Request timed out after ${REQUEST_TIMEOUT_MS / 1_000} seconds: ${path}` };
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
  // Auto-refresh token on 401/403 — but never when the token was explicitly set via set_token.
  // 403 "Request not allowed" from Seismic typically means an expired or wrong-scope token.
  // First: try loading a token saved by the panel process (cross-process panel login).
  // Then: try auto-login with cached credentials (env vars or from a previous panel login).
  if ((res.status === 401 || res.status === 403) && _retry && !isTokenManual()) {
    const savedToken = loadSavedToken();
    if (savedToken && savedToken !== getToken()) {
      setToken(savedToken);
      dbg(`apiFetch: picked up saved token after ${res.status}, retrying`);
      return apiFetch(path, options, false, base);
    }
    const hasCreds = !!(getCachedUsername() || DEFAULT_USERNAME) && !!(getCachedPassword() || DEFAULT_PASSWORD);
    if (hasCreds) {
      const refreshed = await autoLogin();
      if (refreshed) return apiFetch(path, options, false, base);
    }
  }
  if (res.status === 401 || res.status === 403) {
    return {
      status: res.status,
      body: `Authentication failed (HTTP ${res.status} — token expired or missing). The user must sign in via the LiveDoc panel before this action can proceed. Do not call open_form_ui. Do not ask the user for credentials.`,
    };
  }
  const text = await res.text();
  let body: unknown;
  try {
    body = JSON.parse(text);
  } catch {
    body = text;
  }
  return { status: res.status, body };
}

export function isComplex(resp: Record<string, unknown>): boolean {
  const vlData = (resp.variableListData ?? resp.VariableListData) as Array<Record<string, unknown>> | undefined;
  const imgInput = (resp.imageUploadContentInput ?? resp.ImageUploadContentInput) as Record<string, unknown> | undefined;
  const manualSelect = (resp.manualSelectContentInput ?? resp.ManualSelectContentInput) as Record<string, unknown> | undefined;
  const adhoc = (resp.adhocInputs ?? resp.AdhocInputs) as Array<unknown> | undefined;
  return !!(
    vlData?.some(v => v.dataSourceId ?? v.DataSourceId) ||
    (imgInput?.imageUploadContentItems as Array<unknown> | undefined)?.length ||
    (manualSelect?.manualSelectContentItems as Array<unknown> | undefined)?.length ||
    (adhoc?.length ?? 0) > 8
  );
}

export function gf(o: Record<string, unknown>, key: string): unknown {
  return o[key] ?? o[key[0].toUpperCase() + key.slice(1)];
}
