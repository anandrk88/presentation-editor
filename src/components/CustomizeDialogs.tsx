import React, { useState } from "react";
import type { ColorTheme } from "../model/types";
import { FONT_LIST } from "../model/defaults";
import { customFontNames, type CustomFontPair } from "../util/custom";

/** PowerPoint "Customize Colors": one swatch per theme slot, named, savable. */
const SLOTS: { key: keyof ColorTheme; label: string }[] = [
  { key: "dk1", label: "Text / Dark 1" },
  { key: "lt1", label: "Background / Light 1" },
  { key: "dk2", label: "Text / Dark 2" },
  { key: "lt2", label: "Background / Light 2" },
  { key: "accent1", label: "Accent 1" },
  { key: "accent2", label: "Accent 2" },
  { key: "accent3", label: "Accent 3" },
  { key: "accent4", label: "Accent 4" },
  { key: "accent5", label: "Accent 5" },
  { key: "accent6", label: "Accent 6" },
  { key: "hlink", label: "Hyperlink" },
  { key: "folHlink", label: "Followed link" },
];

export function ThemeColorsDialog({ initial, takenNames, onSave, onClose }: {
  initial: ColorTheme;
  takenNames: string[];
  onSave: (t: ColorTheme) => void;
  onClose: () => void;
}) {
  const defaultName = (() => {
    let n = 1;
    while (takenNames.includes(`Custom ${n}`)) n++;
    return `Custom ${n}`;
  })();
  const [name, setName] = useState(initial.name.startsWith("Custom") ? initial.name : defaultName);
  const [colors, setColors] = useState<Record<string, string>>(() => {
    const out: Record<string, string> = {};
    for (const s of SLOTS) out[s.key] = String(initial[s.key]);
    return out;
  });
  const set = (key: string, hex: string) => setColors(c => ({ ...c, [key]: hex.replace(/^#/, "").toUpperCase() }));

  return (
    <div className="modal-overlay" onPointerDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal-card">
        <div className="modal-title">Customize Colors</div>
        <div className="clr-slot-grid">
          {SLOTS.map(s => (
            <label key={s.key} className="clr-slot">
              <input
                type="color"
                value={"#" + colors[s.key]}
                onChange={e => set(s.key, e.target.value)}
              />
              <span>{s.label}</span>
            </label>
          ))}
        </div>
        <div className="modal-row">
          <span className="pane-mini-label">Preview</span>
          <span className="clr-preview">
            {["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map(k => (
              <span key={k} style={{ background: "#" + colors[k] }} />
            ))}
          </span>
        </div>
        <div className="modal-row">
          <span className="pane-mini-label">Name</span>
          <input className="find-input" value={name} onChange={e => setName(e.target.value)} />
        </div>
        <div className="modal-actions">
          <button className="pane-btn" onClick={onClose}>Cancel</button>
          <button
            className="pane-btn primary"
            disabled={!name.trim()}
            onClick={() => onSave({
              ...initial,
              ...Object.fromEntries(SLOTS.map(s => [s.key, colors[s.key]])),
              name: name.trim(),
            } as ColorTheme)}
          >Save &amp; Apply</button>
        </div>
      </div>
    </div>
  );
}

export function FontPairDialog({ takenNames, initialMajor, initialMinor, onSave, onClose }: {
  takenNames: string[];
  initialMajor: string;
  initialMinor: string;
  onSave: (p: CustomFontPair) => void;
  onClose: () => void;
}) {
  const defaultName = (() => {
    let n = 1;
    while (takenNames.includes(`Custom fonts ${n}`)) n++;
    return `Custom fonts ${n}`;
  })();
  const [name, setName] = useState(defaultName);
  const [major, setMajor] = useState(initialMajor);
  const [minor, setMinor] = useState(initialMinor);
  const [google, setGoogle] = useState(false);
  const known = [...new Set([...FONT_LIST, ...customFontNames()])];

  return (
    <div className="modal-overlay" onPointerDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal-card">
        <div className="modal-title">Customize Fonts</div>
        <datalist id="font-suggestions">
          {known.map(f => <option key={f} value={f} />)}
        </datalist>
        <div className="modal-row">
          <span className="pane-mini-label">Heading font</span>
          <input className="find-input" list="font-suggestions" value={major} onChange={e => setMajor(e.target.value)} />
        </div>
        <div className="modal-row">
          <span className="pane-mini-label">Body font</span>
          <input className="find-input" list="font-suggestions" value={minor} onChange={e => setMinor(e.target.value)} />
        </div>
        <div className="modal-row preview-fonts">
          <span style={{ fontFamily: `'${major}', sans-serif`, fontSize: 18 }}>Heading Aa Bb Cc</span>
          <span style={{ fontFamily: `'${minor}', sans-serif`, fontSize: 13 }}>Body text aa bb cc 0123</span>
        </div>
        <label className="pane-check" style={{ margin: "4px 0" }}>
          <input type="checkbox" checked={google} onChange={e => setGoogle(e.target.checked)} />
          Load from Google Fonts (for fonts not installed on this device)
        </label>
        <div className="modal-row">
          <span className="pane-mini-label">Name</span>
          <input className="find-input" value={name} onChange={e => setName(e.target.value)} />
        </div>
        <div className="modal-actions">
          <button className="pane-btn" onClick={onClose}>Cancel</button>
          <button
            className="pane-btn primary"
            disabled={!name.trim() || !major.trim() || !minor.trim()}
            onClick={() => onSave({ name: name.trim(), major: major.trim(), minor: minor.trim(), google: google || undefined })}
          >Save &amp; Apply</button>
        </div>
      </div>
    </div>
  );
}
