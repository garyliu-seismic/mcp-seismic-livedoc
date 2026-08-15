import { useState } from "react";

export function useGenerationPoll({ appRef, tokenRef, setPhase, setErrMsg }) {
  const [result,  setResult]  = useState(null);
  const [pollMsg, setPollMsg] = useState("");

  function resetResult() {
    setResult(null);
    setPollMsg("");
  }

  async function pollUntilDone(app, id) {
    const MAX = 80; // ~4 minutes at 3s interval
    for (let i = 0; i < MAX; i++) {
      await new Promise(r => setTimeout(r, 3000));
      const res = await app.callServerTool({ name: "poll_generation", arguments: { generatedLivedocId: id } });
      const sc = res?.structuredContent;
      const status = sc?.status ?? "Unknown";
      const elapsed = `${(i + 1) * 3}s`;
      if (sc?.outputs?.length) {
        const parts = sc.outputs.map(o => `${o.format}: ${o.status}`).join("  ·  ");
        setPollMsg(`${parts}  (${elapsed})`);
      } else {
        setPollMsg(`Generating… (${elapsed})`);
      }
      if (status === "Completed") return sc ?? {};
      if (status === "Failed") {
        const detail = sc?.outputs ? sc.outputs.map(o => `${o.format}: ${o.errorMessage ?? o.status}`).join("; ") : "";
        throw new Error("Generation failed" + (detail ? ` — ${detail}` : ""));
      }
    }
    throw new Error(`Generation timed out after 4 minutes. Generation ID: ${id} — ask Claude to check its status or download the output.`);
  }

  async function finishGeneration(app, generatedLivedocId, onDone) {
    setPhase("polling");
    const pollResult = await pollUntilDone(app, generatedLivedocId);
    const dls = pollResult.downloads ?? [];
    setResult({ generatedLivedocId, downloadUrls: pollResult.downloadUrls ?? [], downloads: dls });
    setPhase("done");
    app.sendSizeChanged({ width: 520, height: 500 });
    onDone?.(generatedLivedocId, dls);
  }

  async function handleSubmit(getBuildPayload, onDone) {
    const app = appRef.current;
    setPhase("submitting");
    setErrMsg(null);
    try {
      const payload = getBuildPayload();
      const submitRes = await app.callServerTool({
        name: "submit_form",
        arguments: { token: tokenRef.current, payload: JSON.stringify(payload) },
      });
      const sc = submitRes?.structuredContent;
      if (sc?.error) throw new Error(sc.error + (sc.detail ? ` — ${JSON.stringify(sc.detail)}` : ""));
      const generatedLivedocId = sc?.generatedLivedocId;
      if (!generatedLivedocId) throw new Error("No generatedLivedocId in response");

      await finishGeneration(app, generatedLivedocId, onDone);
    } catch (e) {
      setErrMsg(String(e));
      setPhase("error");
    }
  }

  // Resume a form whose generation was already submitted before this panel (re)mounted —
  // e.g. a chat refresh recreated the iframe mid-generation or after it finished. Restores
  // the done/polling view instead of falling back to the blank input form. Returns true if
  // it took over the phase, so the caller can skip its own "ready" transition.
  function resumeFromResult(existingResult, onDone) {
    const app = appRef.current;
    if (!existingResult?.generatedLivedocId) return false;
    if (existingResult.status === "Completed") {
      const dls = existingResult.downloads ?? [];
      setResult({
        generatedLivedocId: existingResult.generatedLivedocId,
        downloadUrls: existingResult.downloadUrls ?? [],
        downloads: dls,
      });
      setPhase("done");
      app.sendSizeChanged({ width: 520, height: 500 });
      onDone?.(existingResult.generatedLivedocId, dls);
      return true;
    }
    setErrMsg(null);
    finishGeneration(app, existingResult.generatedLivedocId, onDone).catch(e => {
      setErrMsg(String(e));
      setPhase("error");
    });
    return true;
  }

  return { result, pollMsg, handleSubmit, resetResult, resumeFromResult };
}
