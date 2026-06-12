import React from "react";
import { store } from "../state/store";
import { useEditorState } from "../state/useStore";
import { Icon } from "./icons";
import { Dropdown } from "./Dropdown";

const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 1.75, 2, 3, 4];

export function StatusBar({ onPresent }: { onPresent: () => void }) {
  const state = useEditorState();
  const { pres, selection } = state;
  const effectiveZoom = state.effectiveZoom;
  const pct = Math.round(effectiveZoom * 100);

  const bump = (dir: 1 | -1) => {
    const cur = effectiveZoom;
    const next = dir === 1
      ? ZOOM_STEPS.find(z => z > cur + 0.001) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1]
      : [...ZOOM_STEPS].reverse().find(z => z < cur - 0.001) ?? ZOOM_STEPS[0];
    store.setState({ zoom: next });
  };

  return (
    <div className="statusbar">
      <button className="sb-btn" title="Start slideshow (F5)" onClick={onPresent}><Icon name="present" size={13} /></button>
      <span className="sb-label">Slide {selection.slideIndex + 1} of {pres.slides.length}</span>
      {state.statusMessage && <span className="sb-msg">{state.statusMessage}</span>}
      <div className="sb-spacer" />
      <span className="sb-label theme-label">Theme: {pres.theme.name}</span>
      <span className="sb-label lang-label">English – United States</span>
      <button className={`sb-btn ${state.zoom === "fit" ? "active" : ""}`} title="Fit to slide" onClick={() => store.setState({ zoom: "fit" })}><Icon name="fit" size={15} /></button>
      <button className={`sb-btn ${state.zoom === "fitw" ? "active" : ""}`} title="Fit to width" onClick={() => store.setState({ zoom: "fitw" })}><Icon name="fitW" size={15} /></button>
      <button className="sb-btn" title="Zoom out" onClick={() => bump(-1)}><Icon name="zoomOut" size={15} /></button>
      <Dropdown className="zoom-dd" panelClassName="up" button={<span className="sb-zoom">Zoom {pct}%</span>} title="Zoom">
        {close => (
          <div className="menu">
            {ZOOM_STEPS.map(z => (
              <button key={z} className="menu-item" onClick={() => { store.setState({ zoom: z }); close(); }}>{Math.round(z * 100)}%</button>
            ))}
          </div>
        )}
      </Dropdown>
      <button className="sb-btn" title="Zoom in" onClick={() => bump(1)}><Icon name="zoomIn" size={15} /></button>
    </div>
  );
}
