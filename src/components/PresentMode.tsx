import React, { useEffect, useLayoutEffect, useRef, useState } from "react";
import { SlideSVG, px } from "../render/SlideView";
import { store } from "../state/store";
import { useEditorState } from "../state/useStore";

const SPEED_MS = { slow: 1000, med: 550, fast: 280 };

export function PresentMode() {
  const state = useEditorState();
  const { pres, presentIndex } = state;
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight });
  const [anim, setAnim] = useState<{ from: number; to: number; t: number } | null>(null);
  const animRef = useRef<number | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);

  const W = px(pres.slideWidth), H = px(pres.slideHeight);
  const scale = Math.min(size.w / W, size.h / H);

  useLayoutEffect(() => {
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight });
    window.addEventListener("resize", onResize);
    rootRef.current?.requestFullscreen?.().catch(() => { });
    return () => {
      window.removeEventListener("resize", onResize);
      if (document.fullscreenElement) document.exitFullscreen().catch(() => { });
    };
  }, []);

  const go = (to: number) => {
    if (to < 0 || to >= pres.slides.length || to === presentIndex || anim) return;
    const trans = pres.slides[Math.max(to, presentIndex) === to ? to : to]?.transition; // incoming slide's transition
    const duration = trans && trans.type !== "none" ? SPEED_MS[trans.speed] : 0;
    if (!duration || to < presentIndex) {
      store.setState({ presentIndex: to });
      return;
    }
    const start = performance.now();
    setAnim({ from: presentIndex, to, t: 0 });
    const tick = (now: number) => {
      const t = Math.min(1, (now - start) / duration);
      setAnim({ from: presentIndex, to, t });
      if (t < 1) animRef.current = requestAnimationFrame(tick);
      else {
        setAnim(null);
        store.setState({ presentIndex: to });
      }
    };
    animRef.current = requestAnimationFrame(tick);
  };

  useEffect(() => () => { if (animRef.current) cancelAnimationFrame(animRef.current); }, []);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") { store.setState({ presenting: false }); }
      else if (e.key === "ArrowRight" || e.key === " " || e.key === "PageDown" || e.key === "Enter") { e.preventDefault(); go(presentIndex + 1); }
      else if (e.key === "ArrowLeft" || e.key === "PageUp" || e.key === "Backspace") { e.preventDefault(); go(presentIndex - 1); }
      else if (e.key === "Home") go(0);
      else if (e.key === "End") go(pres.slides.length - 1);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  const renderSlide = (i: number, style: React.CSSProperties = {}) => (
    <div className="present-slide" style={{ width: W * scale, height: H * scale, ...style }}>
      <SlideSVG pres={pres} slide={pres.slides[i]} media={store.media} width={W * scale} height={H * scale} />
    </div>
  );

  let content: React.ReactNode;
  if (anim) {
    const incoming = pres.slides[anim.to];
    const tr = incoming.transition;
    const t = anim.t;
    const ease = t < 0.5 ? 2 * t * t : 1 - Math.pow(-2 * t + 2, 2) / 2;
    if (tr?.type === "push") {
      const dir = tr.dir ?? "l";
      const dx = dir === "l" ? -1 : dir === "r" ? 1 : 0;
      const dy = dir === "u" ? -1 : dir === "d" ? 1 : 0;
      const sw = W * scale, sh = H * scale;
      content = (
        <>
          {renderSlide(anim.from, { position: "absolute", transform: `translate(${dx * ease * sw}px, ${dy * ease * sh}px)` })}
          {renderSlide(anim.to, { position: "absolute", transform: `translate(${-dx * (1 - ease) * sw}px, ${-dy * (1 - ease) * sh}px)` })}
        </>
      );
    } else {
      content = (
        <>
          {renderSlide(anim.from, { position: "absolute" })}
          {renderSlide(anim.to, { position: "absolute", opacity: ease })}
        </>
      );
    }
  } else {
    content = renderSlide(presentIndex);
  }

  return (
    <div
      className="present-root"
      ref={rootRef}
      onClick={() => go(presentIndex + 1)}
      onContextMenu={e => { e.preventDefault(); go(presentIndex - 1); }}
    >
      <div className="present-center">{content}</div>
      <div className="present-hud">
        {presentIndex + 1} / {pres.slides.length}
        <span className="present-hint">click / → next · ← back · Esc exit</span>
      </div>
    </div>
  );
}
