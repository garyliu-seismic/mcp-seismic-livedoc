export function Lightbox({ images, idx, onClose, onNavigate }) {
  const cur = images[idx];
  const total = images.length;
  const hasPrev = idx > 0;
  const hasNext = idx < total - 1;
  return (
    <div onClick={onClose}
      style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 9999,
        display: "flex", alignItems: "center", justifyContent: "center" }}>
      {total > 1 && (
        <button
          onClick={e => { e.stopPropagation(); onNavigate(Math.max(0, idx - 1)); }}
          disabled={!hasPrev}
          style={{ position: "fixed", left: 16, top: "50%", transform: "translateY(-50%)",
            background: "rgba(255,255,255,0.15)", border: "none", borderRadius: "50%",
            width: 44, height: 44, fontSize: 26, color: "#fff", cursor: hasPrev ? "pointer" : "default",
            opacity: hasPrev ? 1 : 0.25, display: "flex", alignItems: "center", justifyContent: "center" }}>
          ‹
        </button>
      )}
      <div onClick={e => e.stopPropagation()}
        style={{ position: "relative", display: "flex", flexDirection: "column", alignItems: "center",
          maxWidth: "calc(100vw - 100px)", maxHeight: "90vh" }}>
        <img src={cur.url} alt={cur.title}
          style={{ display: "block", maxWidth: "calc(100vw - 100px)", maxHeight: "80vh", borderRadius: 6,
            boxShadow: "0 8px 40px rgba(0,0,0,0.5)", objectFit: "contain" }} />
        <div style={{ textAlign: "center", color: "#ccc", fontSize: 12, marginTop: 8 }}>{cur.title}</div>
        <button onClick={onClose}
          style={{ position: "absolute", top: -12, right: -12, background: "#fff", border: "none",
            borderRadius: "50%", width: 26, height: 26, cursor: "pointer", fontSize: 14,
            fontWeight: 700, lineHeight: "26px", zIndex: 1 }}>✕</button>
      </div>
      {total > 1 && (
        <button
          onClick={e => { e.stopPropagation(); onNavigate(Math.min(total - 1, idx + 1)); }}
          disabled={!hasNext}
          style={{ position: "fixed", right: 16, top: "50%", transform: "translateY(-50%)",
            background: "rgba(255,255,255,0.15)", border: "none", borderRadius: "50%",
            width: 44, height: 44, fontSize: 26, color: "#fff", cursor: hasNext ? "pointer" : "default",
            opacity: hasNext ? 1 : 0.25, display: "flex", alignItems: "center", justifyContent: "center" }}>
          ›
        </button>
      )}
    </div>
  );
}
