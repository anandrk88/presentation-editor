import React, { useEffect, useRef, useState } from "react";
import { SlideSVG, px } from "../render/SlideView";
import { store } from "../state/store";
import { useEditorState } from "../state/useStore";

/**
 * Read-only, touch-friendly slide viewer — for seeing a deck on a phone while
 * on the move. Renders the real SVG renderer fit-to-screen, with swipe / tap /
 * keyboard navigation and a slim progress bar. The current slide is the store's
 * selected slide, so the host can also drive it via the scripting API
 * (selectSlide). Booted by ?view=1 or config.js `viewer: true` (see config.ts).
 */
const SWIPE_PX = 45; // horizontal travel that counts as a slide change

export function SlideViewer() {
  const state = useEditorState();
  const pres = state.pres;
  const slides = pres.slides;
  const idx = Math.min(state.selection.slideIndex, Math.max(0, slides.length - 1));

  const [size, setSize] = useState(() => ({
    w: typeof window !== "undefined" ? window.innerWidth : 1280,
    h: typeof window !== "undefined" ? window.innerHeight : 720,
  }));
  const [barShown, setBarShown] = useState(true);

  useEffect(() => {
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight });
    window.addEventListener("resize", onResize);
    window.addEventListener("orientationchange", onResize);
    return () => {
      window.removeEventListener("resize", onResize);
      window.removeEventListener("orientationchange", onResize);
    };
  }, []);

  const go = (d: number) => store.selectSlide(idx + d);   // selectSlide clamps to range
  const goTo = (n: number) => store.selectSlide(n);

  // keyboard (desktop / connected keyboard)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " ") { e.preventDefault(); go(1); }
      else if (e.key === "ArrowLeft" || e.key === "PageUp") { e.preventDefault(); go(-1); }
      else if (e.key === "Home") goTo(0);
      else if (e.key === "End") goTo(slides.length - 1);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [idx, slides.length]);

  // touch swipe — change slide on a horizontal flick; suppress the synthetic click
  const touch = useRef({ x: 0, y: 0, active: false });
  const suppressClick = useRef(false);
  const onTouchStart = (e: React.TouchEvent) => {
    const t = e.touches[0];
    touch.current = { x: t.clientX, y: t.clientY, active: true };
  };
  const onTouchEnd = (e: React.TouchEvent) => {
    if (!touch.current.active) return;
    touch.current.active = false;
    const t = e.changedTouches[0];
    const dx = t.clientX - touch.current.x, dy = t.clientY - touch.current.y;
    if (Math.abs(dx) > SWIPE_PX && Math.abs(dx) > Math.abs(dy)) {
      suppressClick.current = true;
      go(dx < 0 ? 1 : -1);
    }
  };
  // tap zones: left ~40% = back, right = next (mouse uses this too)
  const onClick = (e: React.MouseEvent) => {
    if (suppressClick.current) { suppressClick.current = false; return; }
    go(e.clientX < size.w * 0.4 ? -1 : 1);
  };

  const W = px(pres.slideWidth), H = px(pres.slideHeight);
  const scale = Math.max(0.01, Math.min(size.w / W, size.h / H));
  const sw = Math.round(W * scale), sh = Math.round(H * scale);

  if (!slides.length) {
    return <div className="viewer-root"><div className="viewer-empty">No slides to show</div></div>;
  }

  return (
    <div
      className="viewer-root"
      onClick={onClick}
      onTouchStart={onTouchStart}
      onTouchEnd={onTouchEnd}
    >
      <div className="viewer-stage">
        <div className="viewer-slide" style={{ width: sw, height: sh }}>
          <SlideSVG pres={pres} slide={slides[idx]} media={store.media} width={sw} height={sh} />
        </div>
      </div>

      <div className={`viewer-bar ${barShown ? "" : "hidden"}`} onClick={e => e.stopPropagation()}>
        <button className="viewer-nav" onClick={() => go(-1)} disabled={idx === 0} aria-label="Previous slide">‹</button>
        <div className="viewer-progress" onClick={e => {
          const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
          goTo(Math.floor(((e.clientX - r.left) / r.width) * slides.length));
        }}>
          <div className="viewer-progress-fill" style={{ width: `${((idx + 1) / slides.length) * 100}%` }} />
        </div>
        <span className="viewer-count">{idx + 1} / {slides.length}</span>
        <button className="viewer-nav" onClick={() => go(1)} disabled={idx === slides.length - 1} aria-label="Next slide">›</button>
        <button className="viewer-nav small" onClick={() => setBarShown(false)} aria-label="Hide controls">⌄</button>
      </div>
      {!barShown && (
        <button className="viewer-show" onClick={e => { e.stopPropagation(); setBarShown(true); }} aria-label="Show controls">⌃</button>
      )}
    </div>
  );
}
