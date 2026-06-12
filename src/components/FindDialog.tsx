import React, { useEffect, useRef, useState } from "react";
import { store } from "../state/store";

/** Floating find & replace panel (Ctrl+F / left-rail search). */
export function FindDialog({ onClose }: { onClose: () => void }) {
  const [find, setFind] = useState("");
  const [replace, setReplace] = useState("");
  const [matchCase, setMatchCase] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const cursor = useRef(0);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => { inputRef.current?.focus(); }, []);
  useEffect(() => { cursor.current = 0; setStatus(null); }, [find, matchCase]);

  const findNext = () => {
    const matches = store.findMatches(find, matchCase);
    if (!matches.length) { setStatus("No matches"); return; }
    const m = matches[cursor.current % matches.length];
    cursor.current = (cursor.current + 1) % matches.length;
    store.selectSlide(m.slideIndex);
    store.selectShapes([m.shapeId]);
    const total = matches.reduce((a, x) => a + x.hits, 0);
    setStatus(`${total} match${total === 1 ? "" : "es"} in ${matches.length} object${matches.length === 1 ? "" : "s"}`);
  };

  const replaceAll = () => {
    const n = store.replaceAll(find, replace, matchCase);
    setStatus(n ? `Replaced ${n} occurrence${n === 1 ? "" : "s"}` : "No matches");
    cursor.current = 0;
  };

  return (
    <div className="find-dialog" onPointerDown={e => e.stopPropagation()}>
      <div className="find-row">
        <input
          ref={inputRef}
          className="find-input"
          placeholder="Find…"
          value={find}
          onChange={e => setFind(e.target.value)}
          onKeyDown={e => {
            if (e.key === "Enter") findNext();
            if (e.key === "Escape") onClose();
            e.stopPropagation();
          }}
        />
        <button className="pane-btn sm" disabled={!find} onClick={findNext}>Find</button>
        <button className="find-close" title="Close" onClick={onClose}>×</button>
      </div>
      <div className="find-row">
        <input
          className="find-input"
          placeholder="Replace with…"
          value={replace}
          onChange={e => setReplace(e.target.value)}
          onKeyDown={e => e.stopPropagation()}
        />
        <button className="pane-btn sm" disabled={!find} onClick={replaceAll}>Replace All</button>
      </div>
      <div className="find-row">
        <label className="pane-check">
          <input type="checkbox" checked={matchCase} onChange={e => setMatchCase(e.target.checked)} />
          Match case
        </label>
        {status && <span className="find-status">{status}</span>}
      </div>
    </div>
  );
}
