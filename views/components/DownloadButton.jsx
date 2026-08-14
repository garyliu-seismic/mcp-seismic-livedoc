import { useState } from "react";

export function DownloadButton({ app, url, fileName, label }) {
  const [state, setState] = useState("idle"); // idle | saving | done | error
  const [msg,   setMsg]   = useState("");
  async function handleClick() {
    if (state === "saving") return;
    setState("saving");
    try {
      const res = await app.callServerTool({ name: "download_output_file", arguments: { url, fileName } });
      const sc = res?.structuredContent;
      if (sc?.error) { setState("error"); setMsg(sc.error); return; }
      setState("done");
      setMsg(sc?.localPath ?? "Saved");
    } catch (e) {
      setState("error");
      setMsg(String(e));
    }
  }
  return (
    <div style={{ marginBottom: 8 }}>
      <button
        onClick={handleClick}
        disabled={state === "saving"}
        style={{
          padding: "8px 18px", background: state === "done" ? "#2e7d32" : "#1a6fb5",
          color: "#fff", border: "none", borderRadius: 6, fontSize: 13,
          fontWeight: 600, cursor: state === "saving" ? "wait" : "pointer",
          opacity: state === "saving" ? 0.7 : 1,
        }}
      >
        {state === "saving" ? "Saving…" : state === "done" ? "✓ Saved" : label}
      </button>
      {msg && (
        <div style={{ fontSize: 11, marginTop: 3, color: state === "error" ? "#c00" : "#555", wordBreak: "break-all" }}>
          {state === "error" ? `Error: ${msg}` : msg}
        </div>
      )}
    </div>
  );
}
