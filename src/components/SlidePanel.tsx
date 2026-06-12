import React, { useEffect, useState } from "react";
import type { Presentation, SlideModel } from "../model/types";
import { SlideSVG, px } from "../render/SlideView";
import { store } from "../state/store";
import { useEditorState } from "../state/useStore";

/**
 * Memoized thumbnail: commits clone only the touched slide, so untouched
 * slides keep object identity and skip re-render — on large decks this is
 * the difference between O(1) and O(slides) SVG work per edit.
 */
const Thumb = React.memo(
  function Thumb({ pres, slide, w, h }: { pres: Presentation; slide: SlideModel; w: number; h: number }) {
    return <SlideSVG pres={pres} slide={slide} media={store.media} width={w} height={h} />;
  },
  (a, b) =>
    a.slide === b.slide && a.w === b.w && a.h === b.h &&
    a.pres.theme === b.pres.theme &&
    a.pres.slideWidth === b.pres.slideWidth && a.pres.slideHeight === b.pres.slideHeight,
);

export function SlidePanel({ onPresentFrom }: { onPresentFrom: (i: number) => void }) {
  const state = useEditorState();
  const { pres, selection } = state;
  const [menu, setMenu] = useState<{ x: number; y: number; slide: number; empty?: boolean } | null>(null);
  const [dropAt, setDropAt] = useState<number | null>(null);

  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(null);
    document.addEventListener("pointerdown", close);
    window.addEventListener("blur", close);
    return () => {
      document.removeEventListener("pointerdown", close);
      window.removeEventListener("blur", close);
    };
  }, [menu]);

  const W = px(pres.slideWidth), H = px(pres.slideHeight);
  const thumbW = 152;
  const thumbH = Math.round((thumbW * H) / W);

  return (
    <>
      <div
        className="slide-list"
        onPointerDown={e => e.stopPropagation()}
        onContextMenu={e => {
          // right-click on the empty area below the thumbnails
          if ((e.target as Element).closest(".thumb-row")) return;
          e.preventDefault();
          setMenu({ x: e.clientX, y: e.clientY, slide: pres.slides.length - 1, empty: true });
        }}
      >
          {pres.slides.map((slide, i) => (
            <div
              key={slide.id}
              className={`thumb-row ${i === selection.slideIndex ? "selected" : ""} ${dropAt === i ? "drop-before" : ""}`}
              tabIndex={0}
              onClick={e => { store.selectSlide(i); e.currentTarget.focus(); }}
              onKeyDown={e => {
                if (e.key === "Delete" || e.key === "Backspace") {
                  e.preventDefault();
                  e.stopPropagation();
                  store.deleteSlide(i);
                }
              }}
              onContextMenu={e => {
                e.preventDefault();
                store.selectSlide(i);
                setMenu({ x: e.clientX, y: e.clientY, slide: i });
              }}
              draggable
              onDragStart={e => {
                e.dataTransfer.setData("text/x-slide-index", String(i));
                e.dataTransfer.effectAllowed = "move";
              }}
              onDragOver={e => {
                e.preventDefault();
                const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
                const before = e.clientY < rect.top + rect.height / 2;
                setDropAt(before ? i : i + 1);
              }}
              onDragLeave={() => setDropAt(null)}
              onDrop={e => {
                e.preventDefault();
                const from = parseInt(e.dataTransfer.getData("text/x-slide-index"), 10);
                let to = dropAt ?? i;
                setDropAt(null);
                if (Number.isNaN(from)) return;
                if (to > from) to -= 1;
                store.moveSlide(from, to);
              }}
            >
              <span className="thumb-num">{i + 1}</span>
              <div className="thumb-frame" style={{ width: thumbW, height: thumbH }}>
                <Thumb pres={pres} slide={slide} w={thumbW} h={thumbH} />
              </div>
            </div>
          ))}
          <div
            className={`thumb-tail ${dropAt === pres.slides.length ? "drop-before" : ""}`}
            onDragOver={e => { e.preventDefault(); setDropAt(pres.slides.length); }}
            onDrop={e => {
              e.preventDefault();
              const from = parseInt(e.dataTransfer.getData("text/x-slide-index"), 10);
              setDropAt(null);
              if (!Number.isNaN(from)) store.moveSlide(from, pres.slides.length - 1);
            }}
          />
      </div>
      {menu && (
        <div className="ctx-menu" style={{ left: menu.x, top: menu.y }} onPointerDown={e => e.stopPropagation()}>
          <button className="menu-item" onClick={() => { store.addSlide("titleContent", menu.slide); setMenu(null); }}>New Slide</button>
          {!menu.empty && (
            <>
              <button className="menu-item" onClick={() => { store.duplicateSlide(menu.slide); setMenu(null); }}>Duplicate Slide</button>
              <button className="menu-item" onClick={() => { store.deleteSlide(menu.slide); setMenu(null); }}>Delete Slide</button>
              <div className="menu-div" />
              <button className="menu-item" disabled={menu.slide === 0} onClick={() => { store.moveSlide(menu.slide, 0); setMenu(null); }}>Move Slide to Beginning</button>
              <button className="menu-item" disabled={menu.slide === pres.slides.length - 1} onClick={() => { store.moveSlide(menu.slide, pres.slides.length - 1); setMenu(null); }}>Move Slide to End</button>
              <div className="menu-div" />
              <button className="menu-item" onClick={() => { setMenu(null); onPresentFrom(menu.slide); }}>Start Slideshow from Here</button>
            </>
          )}
        </div>
      )}
    </>
  );
}
