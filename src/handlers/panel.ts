import { getToken } from "../auth/state.js";
import { openWithDefaultApp, generateToken } from "../utils/os-utils.js";
import { FORM_APP_BASE, FORM_API_BASE } from "../config.js";

export async function handleOpenFormUi(args: { teamSiteId: string; libraryContentVersionId: string; context?: string; prefillValues?: unknown }) {
  // Push the current token to the form server so it never uses a stale value.
  try {
    await fetch(`${FORM_API_BASE}/api/set-token`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ token: getToken() }),
    });
  } catch {
    // Non-fatal — form server may not be running yet; it will fall back to its own env var.
  }

  const token = generateToken();
  const params = new URLSearchParams({
    teamSiteId: args.teamSiteId,
    versionId: args.libraryContentVersionId,
    token,
  });
  if (args.context) {
    params.set("context", Buffer.from(args.context).toString("base64"));
  }
  if (args.prefillValues) {
    params.set("prefill", Buffer.from(JSON.stringify(args.prefillValues)).toString("base64"));
  }
  const url = `${FORM_APP_BASE}/form?${params}`;
  openWithDefaultApp(url);
  return {
    url,
    token,
    message: [
      `The form has been opened in the browser: ${url}`,
      `IMPORTANT: Call get_form_result with token="${token}" after the user submits the form.`,
      `get_form_result polls until the form app posts the result, then returns the generatedLivedocId and outputs.`,
    ].join("\n"),
  };
}

export async function handleGetFormResult(args: { token: string }) {
  const res = await fetch(`${FORM_API_BASE}/api/result/${args.token}`);
  if (res.status === 404) {
    return { error: "Result not ready yet — the form may still be open or generation is in progress. Try again in a moment." };
  }
  if (!res.ok) {
    return { error: `Failed to retrieve result (HTTP ${res.status})` };
  }
  const data = await res.json() as { generatedLivedocId: string; outputs: Array<{ id: string; status: string; format: string; fileName: string }> };
  return {
    generatedLivedocId: data.generatedLivedocId,
    outputs: data.outputs,
    hint: "Generation complete. Call get_generation_download_url with generatedLivedocId and each output's id to get download links.",
  };
}
