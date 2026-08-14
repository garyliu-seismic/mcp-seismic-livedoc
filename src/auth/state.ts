import { DEFAULT_AUTH_TENANT, DEFAULT_USERNAME, DEFAULT_PASSWORD } from "../config.js";

let currentToken = process.env.SEISMIC_API_TOKEN ?? "";
// True once a token was explicitly provided via set_token — disables the silent
// 401-triggered autoLogin() so it can never clobber a hand-picked token with a
// narrower-scoped one obtained from the default credential-flow login.
let tokenIsManual = false;
// Credentials entered via the panel login form — cached in memory so the 401
// auto-refresh path can obtain a fresh token without env vars being set.
let cachedTenant   = DEFAULT_AUTH_TENANT;
let cachedUsername = DEFAULT_USERNAME;
let cachedPassword = DEFAULT_PASSWORD;

export function getToken(): string { return currentToken; }
export function setToken(t: string): void { currentToken = t; }
export function isTokenManual(): boolean { return tokenIsManual; }
export function setTokenManual(v: boolean): void { tokenIsManual = v; }
export function getCachedTenant(): string { return cachedTenant; }
export function setCachedTenant(t: string): void { cachedTenant = t; }
export function getCachedUsername(): string { return cachedUsername; }
export function setCachedUsername(u: string): void { cachedUsername = u; }
export function getCachedPassword(): string { return cachedPassword; }
export function setCachedPassword(p: string): void { cachedPassword = p; }
