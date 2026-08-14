import { useState, useRef, useEffect } from "react";
import { DownloadButton } from "./DownloadButton.jsx";

export function SlidePreview({ images, templateName, downloads, downloadUrls, app, onZoom }) {
  const [idx, setIdx] = useState(0);
  const thumbsRef = useRef(null);
  const total = images.length;
  const cur = images[idx];
  const hasPrev = idx > 0;
  const hasNext = idx < total - 1;

  useEffect(() => {
    const el = thumbsRef.current?.children[idx];
    if (el) el.scrollIntoView({ behavior: "smooth", inline: "nearest", block: "nearest" });
  }, [idx]);

  return (
    <div style={{ borderRadius: 10, overflow: "hidden", border: "1px solid #dde3ea", boxShadow: "0 2px 12px rgba(0,0,0,.08)" }}>
      {/* Header */}
      <div style={{ padding: "12px 16px", background: "#fff", borderBottom: "1px solid #eaeaea", display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15, color: "#1d1d1f" }}>{templateName}</div>
          <div style={{ fontSize: 11, color: "#888", marginTop: 3 }}>LiveDoc generation · slide preview</div>
        </div>
        <div style={{ fontSize: 11, color: "#aaa", marginTop: 2 }}>
          Version <span style={{ fontWeight: 600, color: "#555" }}>Current</span>
        </div>
      </div>
      {/* Download buttons */}
      <div style={{ padding: "10px 14px", background: "#fafafa", borderBottom: "1px solid #efefef", display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        {(downloads.length > 0 ? downloads : []).map((dl, i) => (
          <DownloadButton key={i} app={app} url={dl.url} fileName={dl.fileName} label={`↓ Download ${dl.format.toUpperCase()}`} />
        ))}
        {downloads.length === 0 && downloadUrls.map((url, i) => (
          <DownloadButton key={i} app={app} url={url} fileName={`output-${i + 1}.pptx`} label={`↓ Download ${i + 1}`} />
        ))}
      </div>
      {/* Main slide */}
      <div style={{ background: "#111", position: "relative", display: "flex", alignItems: "center", justifyContent: "center", padding: "16px 44px", minHeight: 180 }}>
        {total > 1 && (
          <button onClick={() => setIdx(i => Math.max(0, i - 1))} disabled={!hasPrev}
            style={{ position: "absolute", left: 6, top: "50%", transform: "translateY(-50%)",
              background: hasPrev ? "rgba(255,255,255,.2)" : "rgba(255,255,255,.05)", border: "none",
              borderRadius: "50%", width: 32, height: 32, fontSize: 22, color: "#fff",
              cursor: hasPrev ? "pointer" : "default", opacity: hasPrev ? 1 : 0.3,
              display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1 }}>‹</button>
        )}
        <img src={cur.url} alt={`Slide ${idx + 1}`}
          onClick={() => onZoom && onZoom(idx)}
          style={{ display: "block", maxWidth: "100%", maxHeight: 280, objectFit: "contain",
            borderRadius: 4, boxShadow: "0 4px 24px rgba(0,0,0,.6)",
            cursor: onZoom ? "zoom-in" : "default" }} />
        {total > 1 && (
          <button onClick={() => setIdx(i => Math.min(total - 1, i + 1))} disabled={!hasNext}
            style={{ position: "absolute", right: 6, top: "50%", transform: "translateY(-50%)",
              background: hasNext ? "rgba(255,255,255,.2)" : "rgba(255,255,255,.05)", border: "none",
              borderRadius: "50%", width: 32, height: 32, fontSize: 22, color: "#fff",
              cursor: hasNext ? "pointer" : "default", opacity: hasNext ? 1 : 0.3,
              display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1 }}>›</button>
        )}
      </div>
      {/* Slide counter */}
      <div style={{ background: "#111", textAlign: "center", paddingBottom: 12, color: "#777", fontSize: 12 }}>
        Slide {idx + 1} of {total}
      </div>
      {/* Thumbnail strip */}
      <div ref={thumbsRef} style={{ display: "flex", overflowX: "auto", gap: 6, padding: "10px 12px", background: "#f5f6f8", borderTop: "1px solid #e4e8ed" }}>
        {images.map((img, i) => (
          <div key={img.index} onClick={() => setIdx(i)} title={`Slide ${i + 1}`}
            style={{ flexShrink: 0, width: 80, borderRadius: 5, overflow: "hidden", cursor: "pointer",
              border: i === idx ? "2.5px solid #0066cc" : "2px solid transparent",
              boxShadow: i === idx ? "0 0 0 2px #b3cdf7" : "0 1px 3px rgba(0,0,0,.1)",
              background: "#fff", transition: "border-color .1s" }}>
            <img src={img.url} alt={`Slide ${i + 1}`}
              style={{ width: "100%", display: "block", aspectRatio: "16/9", objectFit: "cover" }} />
          </div>
        ))}
      </div>
    </div>
  );
}
