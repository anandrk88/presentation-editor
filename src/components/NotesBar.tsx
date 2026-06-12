import React, { useEffect, useState } from "react";
import { store } from "../state/store";
import { useEditorState } from "../state/useStore";

/** Bottom notes strip ("Click to add notes") — commits to slide.notes on blur. */
export function NotesBar() {
  const state = useEditorState();
  const slide = state.pres.slides[state.selection.slideIndex];
  const [text, setText] = useState(slide?.notes ?? "");
  const [collapsed, setCollapsed] = useState(false);

  // re-sync when the active slide changes
  useEffect(() => { setText(slide?.notes ?? ""); }, [slide?.id]);

  const commit = () => {
    if ((slide?.notes ?? "") === text) return;
    const idx = state.selection.slideIndex;
    const slides = store.pres.slides.map((s, i) => (i === idx ? { ...s, notes: text || undefined } : s));
    store.commit({ ...store.pres, slides });
  };

  if (!slide) return null;
  return (
    <div className={`notes-bar ${collapsed ? "collapsed" : ""}`}>
      <button className="notes-toggle" title={collapsed ? "Show notes" : "Hide notes"} onClick={() => setCollapsed(c => !c)}>
        {collapsed ? "▴ Notes" : "▾"}
      </button>
      {!collapsed && (
        <textarea
          className="notes-input"
          placeholder="Click to add notes"
          value={text}
          spellCheck={false}
          onChange={e => setText(e.target.value)}
          onBlur={commit}
        />
      )}
    </div>
  );
}
