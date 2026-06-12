import React, { useState } from "react";
import { importPatternSlide } from "../ooxml/pattern";
import { store } from "../state/store";

/** Paste-a-pattern dialog: SlideBazaar pattern JSON -> new slide. */
export function PatternDialog({ onClose }: { onClose: () => void }) {
  const [text, setText] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const doImport = async () => {
    setBusy(true);
    setError(null);
    try {
      const { slide, media, warnings } = await importPatternSlide(text);
      store.addImportedSlide(slide, media);
      store.setStatus(warnings.length
        ? `Pattern imported — ${warnings.length} item(s) approximated`
        : "Pattern imported");
      setTimeout(() => store.setStatus(null), 4000);
      onClose();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const loadSample = async () => {
    try {
      const res = await fetch("/patterns/growth_report_cover_slide.json");
      setText(await res.text());
      setError(null);
    } catch {
      setError("Sample not found (dev server only)");
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <div className="modal-title">Import Pattern (JSON)</div>
        <div className="modal-note">
          Paste a pattern JSON with <code>elements[].type = "ooxml"</code>. <code>{"{{content.*}}"}</code> placeholders
          are filled with sample values derived from <code>content_schema</code> hints — edit the text on the slide afterwards.
          &nbsp;<button className="link-btn" onClick={loadSample}>Load Growth Report sample</button>
        </div>
        <textarea
          className="modal-textarea"
          placeholder='{"id": "…", "content_schema": {...}, "elements": [{"type": "ooxml", "xml": "<p:sp>…"}]}'
          value={text}
          onChange={e => setText(e.target.value)}
          spellCheck={false}
        />
        {error && <div className="modal-error">{error}</div>}
        <div className="modal-actions">
          <button className="pane-btn primary" disabled={!text.trim() || busy} onClick={doImport}>
            {busy ? "Importing…" : "Import as Slide"}
          </button>
          <button className="pane-btn" onClick={onClose}>Cancel</button>
        </div>
      </div>
    </div>
  );
}
