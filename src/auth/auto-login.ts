import { dbg } from "../utils/debug.js";
import { DEFAULT_AUTH_TENANT, DEFAULT_USERNAME, DEFAULT_PASSWORD } from "../config.js";
import {
  getToken, setToken, setTokenManual,
  getCachedTenant, getCachedUsername, getCachedPassword,
  setCachedTenant, setCachedUsername, setCachedPassword,
} from "./state.js";
import { browserLogin } from "./browser-login.js";
import { saveToken } from "./token-store.js";
import { jwtExpiresAt } from "./jwt.js";

export async function autoLogin(): Promise<boolean> {
  const tenant   = getCachedTenant()   || DEFAULT_AUTH_TENANT;
  const username = getCachedUsername() || DEFAULT_USERNAME;
  const password = getCachedPassword() || DEFAULT_PASSWORD;
  if (!tenant || !username || !password) return false;
  dbg(`autoLogin: attempting tenant=${tenant} user=${username}`);
  try {
    const token = await browserLogin(tenant, username, password);
    setToken(token);
    setTokenManual(false);
    saveToken(token);
    dbg(`autoLogin: success`);
    return true;
  } catch (e) {
    dbg(`autoLogin: failed — ${e}`);
    return false;
  }
}

export async function handleLogin(args: {
  tenant?: string;
  username?: string;
  password?: string;
}): Promise<{ ok: boolean; message: string } | { error: string; detail: unknown }> {
  const tenant   = args.tenant   ?? DEFAULT_AUTH_TENANT;
  const username = args.username ?? DEFAULT_USERNAME;
  const password = args.password ?? DEFAULT_PASSWORD;

  if (!tenant)   return { error: "tenant is required (set AUTH_TENANT env var or pass tenant).", detail: null };
  if (!username) return { error: "username is required.", detail: null };
  if (!password) return { error: "password is required.", detail: null };

  dbg(`handleLogin: attempting login tenant=${tenant} user=${username}`);
  try {
    const token = await browserLogin(tenant, username, password);
    setToken(token);
    setTokenManual(false);
    saveToken(token);
    // Cache credentials so the 401 auto-refresh path can re-login without env vars.
    setCachedTenant(tenant);
    setCachedUsername(username);
    setCachedPassword(password);
    const exp = jwtExpiresAt(getToken());
    dbg(`handleLogin: success — token set, expires=${exp ? new Date(exp).toISOString() : "unknown"}`);
    return { ok: true, message: "Token obtained successfully. All tools are now authenticated." };
  } catch (e) {
    dbg(`handleLogin: failed — ${e}`);
    return { error: String(e), detail: null };
  }
}
