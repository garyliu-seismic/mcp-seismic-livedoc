export function jwtExpiresAt(token: string): number | null {
  try {
    const payload = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
    return typeof payload.exp === "number" ? payload.exp * 1000 : null;
  } catch { return null; }
}

// tenant_fqdn (e.g. "qa01eastasia01.seismic.com") is used to build direct Workspace deep links —
// the LDS API responses only return spaceId/folderId/fileId, no browsable URL.
export function jwtTenantFqdn(token: string): string | null {
  try {
    const payload = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
    return typeof payload.tenant_fqdn === "string" ? payload.tenant_fqdn : null;
  } catch { return null; }
}
