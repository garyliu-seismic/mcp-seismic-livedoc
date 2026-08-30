import { apiFetch, gf } from "../api/client.js";
import { INTEGRATION_BASE_URL } from "../config.js";

// The Integration API's GetListOfMyAssignedProfiles returns every DocCenter/Sales Enablement
// profile the current user can see, with no server-side name/teamSiteId filter — so this
// resolves profileId/profileVersionId by matching client-side, same as the web UCB app's own
// bootstrap flow does (it just gets the ids handed to it instead of looking them up).
export async function handleFindDocCenterProfile(args: {
  profileName: string;
  teamSiteId?: string;
}) {
  const result = await apiFetch("/v2/users/profiles", {}, true, INTEGRATION_BASE_URL);
  if (result.status === 401 || result.status === 403) {
    return {
      error: "profile_lookup_unauthorized",
      message: `The Integration API rejected the request (HTTP ${result.status}) — the current token likely lacks the ` +
        "seismic.self.view/seismic.self.manage scope required by GetListOfMyAssignedProfiles, or isn't authorized for the " +
        "Integration API at all. This is a token/auth issue, not a bug in this tool — a fixed/manual SEISMIC_API_TOKEN " +
        "(as opposed to one obtained via the browser login flow) will not have this scope unless explicitly granted. " +
        "Ask the user to supply a token with self.view/self.manage, or fall back to reading contentProfiles/profileVersionIds " +
        "directly off search_livedoc_templates/search_livedoc_content results.",
    };
  }
  if (result.status === 404) {
    return {
      error: "profile_lookup_not_found",
      message: `Integration API returned 404 at ${INTEGRATION_BASE_URL}/v2/users/profiles — this environment's gateway ` +
        "likely does not route the Integration API at this path (check SEISMIC_BASE_URL / SEISMIC_INTEGRATION_BASE_URL). " +
        "Fall back to reading contentProfiles/profileVersionIds directly off search_livedoc_templates/search_livedoc_content results.",
    };
  }
  if (result.status !== 200) {
    return { error: `Listing profiles failed (HTTP ${result.status})`, detail: result.body };
  }

  const profiles = (result.body as Array<Record<string, unknown>>) ?? [];
  const nameLower = args.profileName.trim().toLowerCase();

  const toResult = (p: Record<string, unknown>) => ({
    profileId: gf(p, "id"),
    profileVersionId: gf(p, "versionId"),
    name: gf(p, "name"),
    teamSiteId: gf(p, "teamSiteId"),
    isDefault: gf(p, "isDefault"),
    isPublished: gf(p, "isPublished"),
  });

  let matches = profiles.filter((p) => String(gf(p, "name") ?? "").toLowerCase() === nameLower);
  if (args.teamSiteId) {
    matches = matches.filter((p) => String(gf(p, "teamSiteId") ?? "") === args.teamSiteId);
  }

  if (matches.length > 0) {
    return { totalCount: matches.length, matches: matches.map(toResult) };
  }

  // No exact match — surface partial-name matches so the caller can disambiguate instead of guessing.
  const suggestions = profiles
    .filter((p) => String(gf(p, "name") ?? "").toLowerCase().includes(nameLower))
    .map(toResult);
  return { totalCount: 0, matches: [], suggestions };
}
