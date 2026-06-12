import React from "react";
import type { ColorRef, ColorTheme, SchemeSlot } from "../model/types";
import { resolveColor } from "../model/defaults";
import { Dropdown } from "./Dropdown";
import { Icon } from "./icons";

const THEME_COLS: SchemeSlot[] = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
const STANDARD = ["C00000", "FF0000", "FFC000", "FFFF00", "92D050", "00B050", "00B0F0", "0070C6", "002060", "7030A0"];

function variantsFor(slot: SchemeSlot, theme: ColorTheme): ColorRef[] {
  const hex = theme[slot];
  const lum = luminance(hex);
  if (lum > 0.75) {
    return [95, 85, 75, 65, 50].map(m => ({ kind: "scheme", slot, lumMod: m }));
  }
  return [
    { kind: "scheme", slot, lumMod: 20, lumOff: 80 },
    { kind: "scheme", slot, lumMod: 40, lumOff: 60 },
    { kind: "scheme", slot, lumMod: 60, lumOff: 40 },
    { kind: "scheme", slot, lumMod: 75 },
    { kind: "scheme", slot, lumMod: 50 },
  ];
}

function luminance(hex: string): number {
  const r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}

export function ColorGrid({ theme, onPick, allowNone, noneLabel, onNone, close }: {
  theme: ColorTheme;
  onPick: (c: ColorRef) => void;
  allowNone?: boolean;
  noneLabel?: string;
  onNone?: () => void;
  close: () => void;
}) {
  const sw = (c: ColorRef, key: React.Key, title?: string) => (
    <button
      key={key}
      type="button"
      className="swatch"
      style={{ background: resolveColor(c, theme) }}
      title={title}
      onClick={() => { onPick(c); close(); }}
    />
  );
  return (
    <div className="color-grid">
      <div className="color-section">Theme colors</div>
      <div className="color-row">
        {THEME_COLS.map(slot => sw({ kind: "scheme", slot }, slot, slot))}
      </div>
      {[0, 1, 2, 3, 4].map(v => (
        <div className="color-row" key={v}>
          {THEME_COLS.map(slot => sw(variantsFor(slot, theme)[v], `${slot}-${v}`))}
        </div>
      ))}
      <div className="color-section">Standard colors</div>
      <div className="color-row">
        {STANDARD.map(hex => sw({ kind: "srgb", hex }, hex, "#" + hex))}
      </div>
      <div className="color-extra">
        {allowNone && (
          <button type="button" className="link-btn" onClick={() => { onNone?.(); close(); }}>
            {noneLabel ?? "No Fill"}
          </button>
        )}
        <label className="link-btn custom-color">
          More colors…
          <input
            type="color"
            onChange={e => { onPick({ kind: "srgb", hex: e.target.value.slice(1).toUpperCase() }); close(); }}
          />
        </label>
      </div>
    </div>
  );
}

/** Normalize a CSS color ("#fff", "#FFCC00", "rgb(a)…") to a 6-digit uppercase hex. */
function cssToHex6(s: string | undefined): string {
  if (!s) return "000000";
  const m6 = /^#([0-9a-f]{6})/i.exec(s);
  if (m6) return m6[1].toUpperCase();
  const m3 = /^#([0-9a-f]{3})$/i.exec(s);
  if (m3) return m3[1].split("").map(c => c + c).join("").toUpperCase();
  const rgb = /^rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(s);
  if (rgb) return [rgb[1], rgb[2], rgb[3]].map(n => (+n).toString(16).padStart(2, "0")).join("").toUpperCase();
  return "000000";
}

/**
 * PowerPoint-style split color button: the icon face applies the current color
 * directly; the side chevron opens the palette.
 */
export function ColorButton({ theme, icon, currentHex, title, onPick, allowNone, noneLabel, onNone, disabled }: {
  theme: ColorTheme;
  icon: string;
  currentHex?: string;
  title: string;
  onPick: (c: ColorRef) => void;
  allowNone?: boolean;
  noneLabel?: string;
  onNone?: () => void;
  disabled?: boolean;
}) {
  return (
    <span className="color-split">
      <button
        type="button"
        className="dd-btn color-main"
        title={title}
        disabled={disabled}
        onClick={() => onPick({ kind: "srgb", hex: cssToHex6(currentHex) })}
      >
        <span className="color-btn-face">
          <Icon name={icon} size={18} />
          <span className="color-bar" style={{ background: currentHex ?? "#000" }} />
        </span>
      </button>
      <Dropdown
        title={`${title} options`}
        disabled={disabled}
        panelClassName="color-panel"
        className="color-chev"
        button={<Icon name="chevDown" size={9} />}
      >
        {close => (
          <ColorGrid theme={theme} onPick={onPick} allowNone={allowNone} noneLabel={noneLabel} onNone={onNone} close={close} />
        )}
      </Dropdown>
    </span>
  );
}
