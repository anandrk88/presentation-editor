import React, { useEffect, useRef, useState } from "react";
import type { ArrowEnd, ArrowKind, ChartKind, ChartShape, ColorRef, ColorTheme, Fill, GradientStop, LineDash, PicShape, SchemeSlot, SpShape, TableShape, TableStyleFamily } from "../model/types";
import { CHART_NAMES, EMU_PER_PX, resolveColor, tableBorderColor, tableCellStyle } from "../model/defaults";
import { isLineGeom } from "../render/geometry";
import { seriesColor } from "../render/GraphicViews";
import { store } from "../state/store";
import { useEditorState } from "../state/useStore";
import { loadImageFile } from "../util/loadImage";
import { ColorGrid } from "./ColorPicker";
import { Dropdown } from "./Dropdown";

const FILL_TYPES = [
  ["solid", "Color fill"],
  ["gradient", "Gradient fill"],
  ["image", "Picture or Texture"],
  ["pattern", "Pattern"],
  ["none", "No fill"],
] as const;

const PATTERN_PRESETS = ["ltDnDiag", "ltUpDiag", "dkDnDiag", "horz", "vert", "cross", "diagCross", "smGrid"];

const ARROW_OPTS: [ArrowKind, string][] = [
  ["none", "None"], ["triangle", "Triangle"], ["arrow", "Open arrow"],
  ["stealth", "Stealth"], ["diamond", "Diamond"], ["oval", "Oval"],
];
/** Set the begin/end arrowhead on a line shape (keeps current size). */
function setLineEnd(id: string, which: "headEnd" | "tailEnd", type: ArrowKind) {
  store.updateShapes([id], s => {
    if (s.kind !== "sp") return s;
    const cur = s.line[which];
    const end: ArrowEnd | undefined = type === "none" ? undefined : { type, w: cur?.w ?? "med", len: cur?.len ?? "med" };
    return { ...s, line: { ...s.line, [which]: end } };
  });
}
/** Set arrowhead size on both ends of a line shape. */
function setArrowSize(id: string, size: "sm" | "med" | "lg") {
  store.updateShapes([id], s => {
    if (s.kind !== "sp") return s;
    const upd = (e?: ArrowEnd): ArrowEnd | undefined => (e ? { ...e, w: size, len: size } : e);
    return { ...s, line: { ...s.line, headEnd: upd(s.line.headEnd), tailEnd: upd(s.line.tailEnd) } };
  });
}

/** Background offers the same fill kinds minus "No fill" (matches PowerPoint). */
const BG_TYPES = FILL_TYPES.filter(([v]) => v !== "none");

// ---------- gradient stops slider (PowerPoint Format-Background style) ----------
const MAX_STOPS = 10;

/** Parse a resolved CSS color ("#RRGGBB" or "rgba(...)") into [r,g,b,a01]. */
function cssToRgba(s: string): [number, number, number, number] {
  const m6 = /^#([0-9a-f]{6})/i.exec(s);
  if (m6) {
    const n = parseInt(m6[1], 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255, 1];
  }
  const m = /rgba?\(([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:[,\s/]+([\d.]+))?\)/.exec(s);
  if (m) return [+m[1], +m[2], +m[3], m[4] !== undefined ? +m[4] : 1];
  return [0, 0, 0, 1];
}

/** Color at fraction t between two stops (flattened to literal sRGB, transparency kept). */
function lerpStopColor(a: ColorRef, b: ColorRef, t: number, theme: ColorTheme): ColorRef {
  const A = cssToRgba(resolveColor(a, theme)), B = cssToRgba(resolveColor(b, theme));
  const mix = A.map((v, i) => v + (B[i] - v) * t);
  const hex = [mix[0], mix[1], mix[2]].map(v => Math.round(v).toString(16).padStart(2, "0")).join("").toUpperCase();
  const alpha = Math.round(mix[3] * 100);
  return alpha < 100 ? { kind: "srgb", hex, alpha } : { kind: "srgb", hex };
}

/** PowerPoint-style brightness of a stop color: +lumOff (lighter) / lumMod-100 (darker). */
function brightnessOf(c: ColorRef): number {
  if (c.lumOff) return Math.round(c.lumOff);
  if (c.lumMod !== undefined && c.lumMod < 100) return Math.round(c.lumMod) - 100;
  return 0;
}

function withBrightness(c: ColorRef, b: number): ColorRef {
  const { lumMod: _m, lumOff: _o, ...rest } = c;
  b = Math.max(-100, Math.min(100, Math.round(b)));
  if (!b) return rest as ColorRef;
  return b > 0
    ? { ...(rest as ColorRef), lumMod: 100 - b, lumOff: b }
    : { ...(rest as ColorRef), lumMod: 100 + b };
}

/** Stop color at a bar position, interpolated from its sorted neighbors. */
function colorAtPos(stops: GradientStop[], pos: number, theme: ColorTheme): ColorRef {
  const sorted = [...stops].sort((x, y) => x.pos - y.pos);
  if (pos <= sorted[0].pos) return { ...sorted[0].color };
  const last = sorted[sorted.length - 1];
  if (pos >= last.pos) return { ...last.color };
  for (let i = 0; i < sorted.length - 1; i++) {
    const a = sorted[i], b = sorted[i + 1];
    if (pos >= a.pos && pos <= b.pos) {
      const t = b.pos === a.pos ? 0 : (pos - a.pos) / (b.pos - a.pos);
      return lerpStopColor(a.color, b.color, t, theme);
    }
  }
  return { ...sorted[0].color };
}

/**
 * PowerPoint-style gradient editor: live gradient bar with draggable stop
 * thumbs (click the bar to add a stop), plus Color / Position / Transparency /
 * Brightness for the selected stop.
 */
function GradientSlider({ stops, theme, onStops }: {
  stops: GradientStop[];
  theme: ColorTheme;
  /** commit=false during drags (preview), true for final/one-shot edits. */
  onStops: (stops: GradientStop[], commit: boolean) => void;
}) {
  const [selIdx, setSelIdx] = useState(0);
  const barRef = useRef<HTMLDivElement>(null);
  const sel = stops[Math.min(selIdx, stops.length - 1)];
  useEffect(() => { if (selIdx >= stops.length) setSelIdx(stops.length - 1); }, [stops.length, selIdx]);

  const sorted = [...stops].map((s, i) => ({ ...s, i })).sort((a, b) => a.pos - b.pos);
  const barCss = `linear-gradient(90deg, ${sorted.map(s => `${resolveColor(s.color, theme)} ${s.pos}%`).join(", ")})`;

  const posFromEvent = (e: { clientX: number }) => {
    const r = barRef.current!.getBoundingClientRect();
    return Math.max(0, Math.min(100, Math.round(((e.clientX - r.left) / r.width) * 100)));
  };

  const startThumbDrag = (e: React.PointerEvent, i: number) => {
    e.stopPropagation();
    e.preventDefault();
    setSelIdx(i);
    const el = e.currentTarget as HTMLElement;
    el.setPointerCapture(e.pointerId);
    const move = (ev: PointerEvent) => {
      const pos = posFromEvent(ev);
      onStops(stops.map((s, k) => (k === i ? { ...s, pos } : s)), false);
    };
    const up = (ev: PointerEvent) => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      const pos = posFromEvent(ev);
      const next = stops.map((s, k) => (k === i ? { ...s, pos } : s)).sort((a, b) => a.pos - b.pos);
      setSelIdx(Math.max(0, next.findIndex(s => s.pos === pos)));
      onStops(next, true);
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
  };

  const addStopAt = (pos: number) => {
    if (stops.length >= MAX_STOPS) return;
    const next = [...stops, { pos, color: colorAtPos(stops, pos, theme) }].sort((a, b) => a.pos - b.pos);
    setSelIdx(next.findIndex(s => s.pos === pos));
    onStops(next, true);
  };

  const setSel = (patch: Partial<GradientStop>) => {
    onStops(stops.map((s, k) => (k === selIdx ? { ...s, ...patch } : s)), true);
  };

  const transparency = sel ? (sel.color.alpha === undefined ? 0 : 100 - sel.color.alpha) : 0;
  const brightness = sel ? brightnessOf(sel.color) : 0;

  return (
    <div className="grad-editor">
      <div className="pane-mini-label" style={{ marginTop: 6 }}>Gradient stops</div>
      <div className="grad-bar-wrap">
        <div className="grad-checker" />
        <div
          ref={barRef}
          className="grad-bar"
          style={{ backgroundImage: barCss }}
          title="Click to add a stop"
          onPointerDown={e => addStopAt(posFromEvent(e))}
        />
        {stops.map((s, i) => (
          <div
            key={i}
            className={`grad-thumb ${i === selIdx ? "sel" : ""}`}
            style={{ left: `${s.pos}%`, background: resolveColor({ ...s.color, alpha: undefined }, theme) }}
            title={`${Math.round(s.pos)}%`}
            onPointerDown={e => startThumbDrag(e, i)}
          />
        ))}
        <div className="grad-stop-btns">
          <button
            className="pane-btn sm" title="Add gradient stop"
            disabled={stops.length >= MAX_STOPS}
            onClick={() => {
              // largest gap midpoint
              const ss = [...stops].sort((a, b) => a.pos - b.pos);
              let bestPos = 50, bestGap = -1;
              const pts = [0, ...ss.map(s => s.pos), 100];
              for (let i = 0; i < pts.length - 1; i++) {
                const gap = pts[i + 1] - pts[i];
                if (gap > bestGap) { bestGap = gap; bestPos = Math.round((pts[i] + pts[i + 1]) / 2); }
              }
              addStopAt(bestPos);
            }}
          >+</button>
          <button
            className="pane-btn sm" title="Remove selected stop"
            disabled={stops.length <= 2}
            onClick={() => {
              const next = stops.filter((_, k) => k !== selIdx);
              setSelIdx(Math.max(0, selIdx - 1));
              onStops(next, true);
            }}
          >−</button>
        </div>
      </div>
      {sel && (
        <>
          <div className="pane-row" style={{ marginTop: 6 }}>
            <span className="pane-mini-label">Color</span>
            <InlineColor
              hex={resolveColor(sel.color, theme)}
              onPick={c => setSel({ color: sel.color.alpha !== undefined ? { ...c, alpha: sel.color.alpha } : c })}
              onNone={() => undefined}
              noneLabel=""
              theme={theme}
            />
            <span className="pane-mini-label">Position</span>
            <input
              type="number" className="pane-num" min={0} max={100}
              value={Math.round(sel.pos)}
              onChange={e => setSel({ pos: Math.max(0, Math.min(100, parseInt(e.target.value, 10) || 0)) })}
            />
            <span className="pane-val">%</span>
          </div>
          <div className="pane-row" style={{ marginTop: 6 }}>
            <span className="pane-mini-label">Transparency</span>
            <input
              type="range" className="pane-slider" min={0} max={100}
              value={transparency}
              onChange={e => {
                const t = parseInt(e.target.value, 10) || 0;
                setSel({ color: { ...sel.color, alpha: t === 0 ? undefined : 100 - t } });
              }}
            />
            <input
              type="number" className="pane-num" min={0} max={100}
              value={transparency}
              onChange={e => {
                const t = Math.max(0, Math.min(100, parseInt(e.target.value, 10) || 0));
                setSel({ color: { ...sel.color, alpha: t === 0 ? undefined : 100 - t } });
              }}
            />
            <span className="pane-val">%</span>
          </div>
          <div className="pane-row" style={{ marginTop: 6 }}>
            <span className="pane-mini-label">Brightness</span>
            <input
              type="range" className="pane-slider" min={-100} max={100}
              value={brightness}
              onChange={e => setSel({ color: withBrightness(sel.color, parseInt(e.target.value, 10) || 0) })}
            />
            <input
              type="number" className="pane-num" min={-100} max={100}
              value={brightness}
              onChange={e => setSel({ color: withBrightness(sel.color, parseInt(e.target.value, 10) || 0) })}
            />
            <span className="pane-val">%</span>
          </div>
        </>
      )}
    </div>
  );
}

// ---------- table design (PowerPoint Table Design / Layout parity) ----------
const TBL_FAMILIES: { key: TableStyleFamily; label: string }[] = [
  { key: "light", label: "Light" },
  { key: "medium", label: "Medium" },
  { key: "dark", label: "Dark" },
];
const TBL_ACCENTS: (SchemeSlot | "none")[] = ["none", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

function TableStyleChip({ table, family, accent, theme, active, onPick }: {
  table: TableShape;
  family: TableStyleFamily;
  accent: SchemeSlot | "none";
  theme: ColorTheme;
  active: boolean;
  onPick: () => void;
}) {
  // mini 5×4 preview honoring the table's current style-option flags
  const fake = {
    ...table, styleFamily: family, accent,
    cells: Array.from({ length: 4 }, () => new Array(5).fill({})),
  } as TableShape;
  const border = resolveColor(tableBorderColor({ styleFamily: family, accent, borderColor: undefined }), theme);
  const cw = 6, ch = 5;
  return (
    <button
      type="button"
      className={`tbl-chip ${active ? "active" : ""}`}
      onClick={onPick}
      title={`${family[0].toUpperCase()}${family.slice(1)} — ${accent === "none" ? "no accent" : accent}`}
    >
      <svg viewBox={`0 0 ${cw * 5 + 1} ${ch * 4 + 1}`} shapeRendering="crispEdges">
        {Array.from({ length: 4 }, (_, r) => Array.from({ length: 5 }, (_, c) => {
          const st = tableCellStyle(fake, r, c);
          const clr = st.fill.kind === "solid" ? resolveColor(st.fill.color, theme) : "#FFFFFF";
          return <rect key={`${r}-${c}`} x={c * cw + 0.5} y={r * ch + 0.5} width={cw} height={ch} fill={clr} stroke={border} strokeWidth={0.6} />;
        }))}
      </svg>
    </button>
  );
}

function TablePane({ table, theme, editingCell, tableSel, cm, updateNum }: {
  table: TableShape;
  theme: ColorTheme;
  editingCell: { row: number; col: number } | null;
  tableSel: { id: string; r0: number; c0: number; r1: number; c1: number } | null;
  cm: (emu: number) => number;
  updateNum: (field: "x" | "y" | "w" | "h" | "rot", v: number) => void;
}) {
  const id = table.id;
  const apply = (patch: Partial<TableShape>) => store.applyTableStyle(id, patch);
  const set = (patch: Partial<TableShape>) => store.updateShapes([id], s => (s.kind === "table" ? { ...s, ...patch } : s));
  // active target: cell range > the cell being text-edited > last cell
  const sel = tableSel && tableSel.id === id ? tableSel : null;
  const r0 = sel ? sel.r0 : editingCell?.row ?? table.cells.length - 1;
  const c0 = sel ? sel.c0 : editingCell?.col ?? (table.cells[0]?.length ?? 1) - 1;
  const r1 = sel ? sel.r1 : r0;
  const c1 = sel ? sel.c1 : c0;
  const isRange = !!sel && (r0 !== r1 || c0 !== c1);
  const anchorCell = table.cells[r0]?.[c0];
  const canSplit = !!anchorCell && ((anchorCell.gridSpan ?? 1) > 1 || (anchorCell.rowSpan ?? 1) > 1);
  const targeted = sel || editingCell;
  const cellLabel = isRange ? `R${r0 + 1}:R${r1 + 1} × C${c0 + 1}:C${c1 + 1}` : `R${r0 + 1}C${c0 + 1}`;
  const family = table.styleFamily ?? "medium";
  const accent = table.accent ?? "accent1";
  const row = r0, col = c0;
  const cellFill = anchorCell?.fill;
  const OPTS: { key: "firstRow" | "totalRow" | "bandRow" | "firstCol" | "lastCol" | "bandCol"; label: string }[] = [
    { key: "firstRow", label: "Header row" }, { key: "firstCol", label: "First column" },
    { key: "totalRow", label: "Total row" }, { key: "lastCol", label: "Last column" },
    { key: "bandRow", label: "Banded rows" }, { key: "bandCol", label: "Banded columns" },
  ];
  return (
    <>
      <div className="pane-title">Table settings</div>
      <div className="pane-section">
        <div className="pane-label">Style options</div>
        <div className="tbl-opts">
          {OPTS.map(o => (
            <label key={o.key} className="pane-check">
              <input type="checkbox" checked={!!table[o.key]} onChange={e => apply({ [o.key]: e.target.checked } as Partial<TableShape>)} />
              {o.label}
            </label>
          ))}
        </div>
      </div>
      <div className="pane-section">
        <div className="pane-label">Table styles</div>
        {TBL_FAMILIES.map(f => (
          <div className="tbl-style-row" key={f.key}>
            {TBL_ACCENTS.map(a => (
              <TableStyleChip
                key={a} table={table} family={f.key} accent={a} theme={theme}
                active={family === f.key && accent === a}
                onPick={() => apply({ styleFamily: f.key, accent: a })}
              />
            ))}
          </div>
        ))}
      </div>
      <div className="pane-section">
        <div className="pane-label">Borders</div>
        <div className="pane-row">
          <select className="pane-select" value={table.borderMode ?? "all"} onChange={e => set({ borderMode: e.target.value as TableShape["borderMode"] })}>
            <option value="all">All borders</option>
            <option value="outside">Outside only</option>
            <option value="none">No borders</option>
          </select>
          <InlineColor
            hex={resolveColor(tableBorderColor(table), theme)}
            onPick={c => set({ borderColor: c })}
            onNone={() => set({ borderColor: undefined })}
            noneLabel="Automatic"
            theme={theme}
          />
        </div>
      </div>
      <div className="pane-section">
        <div className="pane-label">Merge cells</div>
        <div className="pane-row">
          <button className="pane-btn sm" disabled={!isRange} title="Merge the selected range" onClick={() => store.mergeTableCells(id, r0, c0, r1, c1)}>Merge</button>
          <button className="pane-btn sm" disabled={!canSplit} title="Split the merged cell back to its grid" onClick={() => store.splitTableCell(id, r0, c0)}>Split</button>
        </div>
        {!sel && <span className="pane-mini-label">drag across cells on the slide to select a range</span>}
      </div>
      <div className="pane-section">
        <div className="pane-label">Cell shading{targeted ? ` — ${cellLabel}` : ""}</div>
        <div className="pane-row">
          <InlineColor
            hex={cellFill?.kind === "solid" ? resolveColor(cellFill.color, theme) : "transparent"}
            onPick={c => store.setCellRangeFill(id, r0, c0, r1, c1, { kind: "solid", color: c })}
            onNone={() => store.setCellRangeFill(id, r0, c0, r1, c1, undefined)}
            noneLabel="From style"
            theme={theme}
          />
          {!targeted && <span className="pane-mini-label">select cells to shade them</span>}
        </div>
      </div>
      <div className="pane-section">
        <div className="pane-label">Rows &amp; columns</div>
        <div className="pane-row">
          <button className="pane-btn sm" title="Insert row above" onClick={() => store.insertTableRow(id, row, true)}>Row ↑</button>
          <button className="pane-btn sm" title="Insert row below" onClick={() => store.insertTableRow(id, row, false)}>Row ↓</button>
          <button className="pane-btn sm" title="Delete row" disabled={table.cells.length <= 1} onClick={() => store.deleteTableRow(id, row)}>− Row</button>
        </div>
        <div className="pane-row" style={{ marginTop: 4 }}>
          <button className="pane-btn sm" title="Insert column left" onClick={() => store.insertTableCol(id, col, true)}>Col ←</button>
          <button className="pane-btn sm" title="Insert column right" onClick={() => store.insertTableCol(id, col, false)}>Col →</button>
          <button className="pane-btn sm" title="Delete column" disabled={(table.cells[0]?.length ?? 1) <= 1} onClick={() => store.deleteTableCol(id, col)}>− Col</button>
        </div>
        <div className="pane-row" style={{ marginTop: 4 }}>
          <button className="pane-btn sm" onClick={() => store.distributeTableRows(id)}>Distribute rows</button>
          <button className="pane-btn sm" onClick={() => store.distributeTableCols(id)}>Distribute cols</button>
        </div>
      </div>
      <div className="pane-section">
        <div className="pane-label">Cell text{targeted ? ` — ${cellLabel}` : " (all cells)"}</div>
        <div className="pane-row">
          <button className="pane-btn sm" title="Align top" onClick={() => targeted ? store.setCellRangeAnchor(id, r0, c0, r1, c1, "t") : store.setTableCellsAnchor(id, "t")}>⤒</button>
          <button className="pane-btn sm" title="Align middle" onClick={() => targeted ? store.setCellRangeAnchor(id, r0, c0, r1, c1, "ctr") : store.setTableCellsAnchor(id, "ctr")}>≡</button>
          <button className="pane-btn sm" title="Align bottom" onClick={() => targeted ? store.setCellRangeAnchor(id, r0, c0, r1, c1, "b") : store.setTableCellsAnchor(id, "b")}>⤓</button>
          <select
            className="pane-select"
            value=""
            onChange={e => {
              const m = e.target.value;
              if (!m) return;
              const ins: [number, number, number, number] =
                m === "narrow" ? [45720, 22860, 45720, 22860]
                : m === "wide" ? [182880, 91440, 182880, 91440]
                : [91440, 45720, 91440, 45720];
              store.setTableCellsInsets(id, ins);
              e.target.value = "";
            }}
            title="Cell margins"
          >
            <option value="" disabled>Cell margins…</option>
            <option value="normal">Normal</option>
            <option value="narrow">Narrow</option>
            <option value="wide">Wide</option>
          </select>
        </div>
      </div>
      <SizePos shape={table} cm={cm} updateNum={updateNum} />
    </>
  );
}

/** Collapsible settings group (remembers open/closed per id in localStorage). */
function Acc({ id, title, defaultOpen, children }: { id: string; title: string; defaultOpen?: boolean; children: React.ReactNode }) {
  const key = "pe.acc." + id;
  const [open, setOpen] = useState(() => {
    try { const v = localStorage.getItem(key); return v === null ? !!defaultOpen : v === "1"; } catch { return !!defaultOpen; }
  });
  const toggle = () => setOpen(o => { try { localStorage.setItem(key, o ? "0" : "1"); } catch { /* quota */ } return !o; });
  return (
    <div className={`acc ${open ? "open" : ""}`}>
      <button type="button" className="acc-head" onClick={toggle}>
        <span className="acc-chev">{open ? "▾" : "▸"}</span>
        <span>{title}</span>
      </button>
      {open && <div className="acc-body">{children}</div>}
    </div>
  );
}

/** Set one pie-slice color override (null = automatic); collapses to undefined when all automatic. */
function setSlice(sh: ChartShape, i: number, c: ColorRef | null): (ColorRef | null)[] | undefined {
  const arr = [...(sh.pointColors ?? sh.categories.map(() => null as ColorRef | null))];
  while (arr.length <= i) arr.push(null);
  arr[i] = c;
  return arr.some(Boolean) ? arr : undefined;
}

type Pane = "slide" | "shape" | "image" | "table" | "chart" | null;

/** Right settings: 40px icon rail + expandable pane, auto-activated by selection. */
export function RightPanel() {
  const state = useEditorState();
  const [open, setOpen] = useState<Pane>("slide");
  const { pres, selection } = state;
  const theme = pres.theme;

  const sel = store.selectedShapes;
  const sp = sel.length === 1 && sel[0].kind === "sp" ? (sel[0] as SpShape) : null;
  const pic = sel.length === 1 && sel[0].kind === "pic" ? sel[0] : null;
  const table = sel.length === 1 && sel[0].kind === "table" ? sel[0] : null;
  const chart = sel.length === 1 && sel[0].kind === "chart" ? sel[0] : null;
  const available: Pane = pic ? "image" : table ? "table" : chart ? "chart" : sel.length ? "shape" : "slide";

  // follow selection like asc_onFocusObject -> RightMenu
  useEffect(() => {
    if (open !== null) setOpen(available);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [available]);

  const slide = pres.slides[selection.slideIndex];
  const shape = sp ?? pic;

  const updateNum = (field: "x" | "y" | "w" | "h" | "rot", v: number) => {
    if (!shape || Number.isNaN(v)) return;
    store.updateShapes([shape.id], s =>
      field === "rot" ? { ...s, rot: ((v % 360) + 360) % 360 } : { ...s, [field]: Math.round(v * EMU_PER_PX) });
  };

  const cm = (emu: number) => Math.round((emu / EMU_PER_PX) * 100) / 100;

  return (
    <div className="right-area">
      {open && (
        <div className="right-pane">
          {open === "slide" && (
            <>
              <div className="pane-title">Slide settings</div>
              <FillControls
                label="Background"
                fill={slide.background ?? { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } }}
                setFill={f => store.updateSlideBg(f)}
                setFillLive={f => store.updateSlideBg(f, { historic: false })}
                theme={theme}
                types={BG_TYPES}
              />
              <div className="pane-section">
                <div className="pane-row">
                  <button className="pane-btn" onClick={() => store.applyBgToAll()}>Apply to All Slides</button>
                  <button
                    className="pane-btn sm"
                    disabled={!slide.background}
                    title="Inherit the default (master) background"
                    onClick={() => store.updateSlideBg(undefined)}
                  >Reset</button>
                </div>
              </div>
            </>
          )}
          {open === "shape" && !shape && sel.length > 1 && (
            <>
              <div className="pane-title">{store.selectionHasGroup ? "Group settings" : "Selection"}</div>
              <div className="pane-section">
                <div className="pane-label">Group</div>
                <div className="pane-row">
                  <button className="pane-btn sm" disabled={sel.length < 2} onClick={() => store.groupSelection()}>Group</button>
                  <button className="pane-btn sm" disabled={!store.selectionHasGroup} onClick={() => store.ungroupSelection()}>Ungroup</button>
                </div>
              </div>
              <div className="pane-section">
                <div className="pane-label">Rotation</div>
                <div className="pane-row">
                  <button className="pane-btn sm" title="Rotate 90° counterclockwise" onClick={() => store.rotateSelected(-90)}>⟲ 90°</button>
                  <button className="pane-btn sm" title="Rotate 90° clockwise" onClick={() => store.rotateSelected(90)}>⟳ 90°</button>
                  <button className="pane-btn sm" title="Flip horizontally" onClick={() => store.flipSelected("h")}>⇋</button>
                  <button className="pane-btn sm" title="Flip vertically" onClick={() => store.flipSelected("v")}>⇅</button>
                </div>
              </div>
            </>
          )}
          {open === "shape" && shape && (
            <>
              <div className="pane-title">Shape settings</div>
              {sp && (
                <>
                  <FillEditor sp={sp} theme={theme} />
                  <div className="pane-section">
                    <div className="pane-label">Stroke</div>
                    <div className="pane-row">
                      <InlineColor
                        hex={sp.line.fill.kind === "solid" ? resolveColor(sp.line.fill.color, theme) : "transparent"}
                        onPick={c => store.updateShapes([sp.id], s => s.kind === "sp" ? { ...s, line: { ...s.line, fill: { kind: "solid", color: c } } } : s)}
                        onNone={() => store.updateShapes([sp.id], s => s.kind === "sp" ? { ...s, line: { ...s.line, fill: { kind: "none" } } } : s)}
                        noneLabel="No Outline"
                        theme={theme}
                      />
                      <select
                        className="pane-select"
                        title="Weight"
                        value={String(sp.line.widthPt)}
                        onChange={e => store.updateShapes([sp.id], s => s.kind === "sp" ? { ...s, line: { ...s.line, widthPt: parseFloat(e.target.value) } } : s)}
                      >
                        {[0.5, 1, 1.5, 2, 2.5, 3, 4.5, 6].map(w => <option key={w} value={String(w)}>{w} pt</option>)}
                      </select>
                    </div>
                    <div className="pane-row" style={{ marginTop: 6 }}>
                      <span className="pane-mini-label">Dashes</span>
                      <select
                        className="pane-select wide"
                        value={sp.line.dash ?? "solid"}
                        onChange={e => store.updateShapes([sp.id], s => s.kind === "sp" ? { ...s, line: { ...s.line, dash: e.target.value === "solid" ? undefined : e.target.value as LineDash } } : s)}
                      >
                        <option value="solid">───── Solid</option>
                        <option value="dot">· · · · · Dotted</option>
                        <option value="sysDot">‧‧‧‧‧ Fine dots</option>
                        <option value="dash">– – – Dashed</option>
                        <option value="dashDot">–·–·– Dash dot</option>
                        <option value="lgDash">—  —  — Long dash</option>
                        <option value="lgDashDot">—·—·— Long dash dot</option>
                      </select>
                    </div>
                  </div>
                  {isLineGeom(sp.geom) && (
                    <div className="pane-section">
                      <div className="pane-label">Arrows</div>
                      <div className="pane-row">
                        <span className="pane-mini-label">Begin</span>
                        <select className="pane-select" value={sp.line.headEnd?.type ?? "none"}
                          onChange={e => setLineEnd(sp.id, "headEnd", e.target.value as ArrowKind)}>
                          {ARROW_OPTS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                        </select>
                        <span className="pane-mini-label">End</span>
                        <select className="pane-select" value={sp.line.tailEnd?.type ?? "none"}
                          onChange={e => setLineEnd(sp.id, "tailEnd", e.target.value as ArrowKind)}>
                          {ARROW_OPTS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                        </select>
                      </div>
                      <div className="pane-row" style={{ marginTop: 6 }}>
                        <span className="pane-mini-label">Size</span>
                        <select className="pane-select wide" value={sp.line.headEnd?.len ?? sp.line.tailEnd?.len ?? "med"}
                          onChange={e => setArrowSize(sp.id, e.target.value as "sm" | "med" | "lg")}>
                          <option value="sm">Small</option>
                          <option value="med">Medium</option>
                          <option value="lg">Large</option>
                        </select>
                      </div>
                    </div>
                  )}
                </>
              )}
              {sp && sp.fill.kind === "solid" && (
                <div className="pane-section">
                  <div className="pane-label">Opacity</div>
                  <div className="pane-row">
                    <input
                      type="range" min={0} max={100}
                      value={sp.fill.color.alpha ?? 100}
                      onChange={e => {
                        const v = parseInt(e.target.value, 10);
                        store.updateShapes([sp.id], s => s.kind === "sp" && s.fill.kind === "solid"
                          ? { ...s, fill: { kind: "solid", color: { ...s.fill.color, alpha: v >= 100 ? undefined : v } } }
                          : s);
                      }}
                    />
                    <span className="pane-val">{sp.fill.color.alpha ?? 100}%</span>
                  </div>
                </div>
              )}
              <SizePos shape={shape} cm={cm} updateNum={updateNum} />
              <div className="pane-section">
                <div className="pane-label">Rotation</div>
                <div className="pane-row">
                  <button className="pane-btn sm" title="Rotate 90° counterclockwise" onClick={() => store.updateShapes([shape.id], s => ({ ...s, rot: ((s.rot - 90) % 360 + 360) % 360 }))}>⟲ 90°</button>
                  <button className="pane-btn sm" title="Rotate 90° clockwise" onClick={() => store.updateShapes([shape.id], s => ({ ...s, rot: (s.rot + 90) % 360 }))}>⟳ 90°</button>
                  <button className="pane-btn sm" title="Flip horizontally" onClick={() => store.updateShapes([shape.id], s => ({ ...s, flipH: !s.flipH || undefined }))}>⇋</button>
                  <button className="pane-btn sm" title="Flip vertically" onClick={() => store.updateShapes([shape.id], s => ({ ...s, flipV: !s.flipV || undefined }))}>⇅</button>
                </div>
              </div>
            </>
          )}
          {open === "image" && pic && (
            <>
              <div className="pane-title">{store.media.get(pic.mediaId)?.mime === "image/svg+xml" ? "Graphic settings" : "Image settings"}</div>
              {store.media.get(pic.mediaId)?.mime === "image/svg+xml" && (
                <div className="pane-section">
                  <div className="pane-label">Graphic color</div>
                  <div className="pane-row">
                    <InlineColor
                      hex={pic.svgTint ? resolveColor(pic.svgTint, theme) : "transparent"}
                      onPick={c => store.updateShapes([pic.id], s => s.kind === "pic" ? { ...s, svgTint: c } : s)}
                      onNone={() => store.updateShapes([pic.id], s => s.kind === "pic" ? { ...s, svgTint: undefined } : s)}
                      noneLabel="Original colors"
                      theme={theme}
                    />
                    <span className="pane-mini-label">{pic.svgTint ? "Recolored" : "Original"}</span>
                  </div>
                </div>
              )}
              <div className="pane-section">
                <div className="pane-label">Corners</div>
                <div className="pane-row">
                  <select
                    className="pane-select wide"
                    value={pic.geom ?? "rect"}
                    onChange={e => {
                      const g = e.target.value;
                      store.updateShapes([pic.id], s => s.kind === "pic"
                        ? { ...s, geom: g === "rect" ? undefined : g, adj: g === "roundRect" ? { adj: 16667 } : undefined }
                        : s);
                    }}
                  >
                    <option value="rect">Square</option>
                    <option value="roundRect">Rounded</option>
                    <option value="ellipse">Circle / Ellipse</option>
                  </select>
                </div>
                {pic.geom === "roundRect" && (() => {
                  const radius = Math.round(((pic as PicShape).adj?.adj ?? 16667) / 1000);
                  const setRadius = (v: number) => {
                    if (!Number.isFinite(v)) return;
                    const clamped = Math.max(0, Math.min(50, Math.round(v)));
                    store.updateShapes([pic.id], s => s.kind === "pic" ? { ...s, adj: { adj: clamped * 1000 } } : s);
                  };
                  return (
                    <div className="pane-row" style={{ marginTop: 6 }}>
                      <span className="pane-mini-label">Radius</span>
                      <input
                        type="range" min={0} max={50} className="pane-slider"
                        value={radius}
                        onChange={e => setRadius(parseInt(e.target.value, 10))}
                      />
                      <input
                        type="number" min={0} max={50} className="pane-num"
                        value={radius}
                        onChange={e => setRadius(parseInt(e.target.value, 10))}
                      />
                      <span className="pane-val">%</span>
                    </div>
                  );
                })()}
              </div>
              <div className="pane-section">
                <div className="pane-label">Crop</div>
                <div className="pane-row">
                  <button
                    className={`pane-btn ${state.croppingId === pic.id ? "primary" : ""}`}
                    title="Crop with handles on the canvas (Esc or click outside to finish)"
                    onClick={() => store.setState({ croppingId: state.croppingId === pic.id ? null : pic.id })}
                  >{state.croppingId === pic.id ? "Done Cropping" : "Crop"}</button>
                  <button
                    className="pane-btn sm"
                    disabled={!pic.srcRect}
                    onClick={() => store.updateShapes([pic.id], s => s.kind === "pic" ? { ...s, srcRect: undefined } : s)}
                  >Reset</button>
                </div>
                <div className="pane-label" style={{ marginTop: 8 }}>Crop (% per edge)</div>
                <div className="pane-grid">
                  {(["l", "t", "r", "b"] as const).map(edge => (
                    <label key={edge}>{edge.toUpperCase()} <input
                      type="number" min={0} max={90}
                      value={Math.round((pic.srcRect?.[edge] ?? 0) * 100)}
                      onChange={e => {
                        const v = Math.max(0, Math.min(90, parseInt(e.target.value, 10) || 0)) / 100;
                        store.updateShapes([pic.id], s => {
                          if (s.kind !== "pic") return s;
                          const sr = { l: 0, t: 0, r: 0, b: 0, ...s.srcRect, [edge]: v };
                          const any = sr.l || sr.t || sr.r || sr.b;
                          return { ...s, srcRect: any ? sr : undefined };
                        });
                      }}
                    /></label>
                  ))}
                </div>
              </div>
              <SizePos shape={pic} cm={cm} updateNum={updateNum} />
              <div className="pane-section">
                <div className="pane-label">Rotation</div>
                <div className="pane-row">
                  <button className="pane-btn sm" onClick={() => store.updateShapes([pic.id], s => ({ ...s, rot: ((s.rot - 90) % 360 + 360) % 360 }))}>⟲ 90°</button>
                  <button className="pane-btn sm" onClick={() => store.updateShapes([pic.id], s => ({ ...s, rot: (s.rot + 90) % 360 }))}>⟳ 90°</button>
                  <button className="pane-btn sm" onClick={() => store.updateShapes([pic.id], s => ({ ...s, flipH: !s.flipH || undefined }))}>⇋</button>
                  <button className="pane-btn sm" onClick={() => store.updateShapes([pic.id], s => ({ ...s, flipV: !s.flipV || undefined }))}>⇅</button>
                </div>
              </div>
            </>
          )}
          {open === "table" && table && (
            <TablePane table={table} theme={theme} editingCell={state.editingShapeId === table.id ? state.editingCell : null} tableSel={state.tableSel} cm={cm} updateNum={updateNum} />
          )}
          {open === "chart" && chart && (
            <>
              <div className="pane-title">Chart settings</div>
              <Acc id="chart-type" title="Type &amp; options" defaultOpen>
              <div className="pane-section">
                <div className="pane-label">Chart type</div>
                <select
                  className="pane-select wide"
                  value={chart.chart}
                  onChange={e => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, chart: e.target.value as ChartKind } : s)}
                >
                  {(Object.keys(CHART_NAMES) as ChartKind[]).map(k => (
                    <option key={k} value={k}>{CHART_NAMES[k]}</option>
                  ))}
                </select>
              </div>
              {["column", "bar", "line", "area"].includes(chart.chart) && (
                <div className="pane-section">
                  <div className="pane-label">Grouping</div>
                  <select
                    className="pane-select wide"
                    value={chart.grouping ?? "clustered"}
                    onChange={e => store.updateShapes([chart.id], s => s.kind === "chart"
                      ? { ...s, grouping: e.target.value === "clustered" ? undefined : e.target.value as ChartShape["grouping"] }
                      : s)}
                  >
                    <option value="clustered">Clustered</option>
                    <option value="stacked">Stacked</option>
                    <option value="percentStacked">100% Stacked</option>
                  </select>
                </div>
              )}
              {chart.chart === "radar" && (
                <div className="pane-section">
                  <div className="pane-label">Radar style</div>
                  <select
                    className="pane-select wide"
                    value={chart.radarStyle ?? "standard"}
                    onChange={e => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, radarStyle: e.target.value as ChartShape["radarStyle"] } : s)}
                  >
                    <option value="standard">Standard</option>
                    <option value="marker">With markers</option>
                    <option value="filled">Filled</option>
                  </select>
                </div>
              )}
              <div className="pane-section">
                <label className="pane-check">
                  <input type="checkbox" checked={chart.legend} onChange={e => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, legend: e.target.checked } : s)} />
                  Show legend
                </label>
                {(chart.chart === "line" || chart.chart === "scatter") && (
                  <>
                    <label className="pane-check">
                      <input type="checkbox" checked={chart.marker ?? true} onChange={e => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, marker: e.target.checked } : s)} />
                      Point markers
                    </label>
                    <label className="pane-check">
                      <input type="checkbox" checked={!!chart.smooth} onChange={e => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, smooth: e.target.checked ? true : (s.chart === "scatter" ? false : undefined) } : s)} />
                      Smooth lines
                    </label>
                  </>
                )}
              </div>
              </Acc>
              <Acc id="chart-elements" title="Chart elements">
              <div className="pane-section">
                <ChartElementsMenu chart={chart} />
              </div>
              </Acc>
              <Acc id="chart-area" title="Chart area &amp; gridlines">
              <div className="pane-section">
                <div className="pane-row">
                  <span className="pane-mini-label">Fill</span>
                  <InlineColor
                    hex={chart.chartFill?.kind === "solid" ? resolveColor(chart.chartFill.color, theme) : chart.chartFill?.kind === "none" ? "transparent" : "#FFFFFF"}
                    onPick={c => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, chartFill: { kind: "solid", color: c } } : s)}
                    onNone={() => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, chartFill: { kind: "none" } } : s)}
                    noneLabel="No Fill (transparent)"
                    theme={theme}
                  />
                  <span className="pane-mini-label">Border</span>
                  <InlineColor
                    hex={chart.chartBorder && chart.chartBorder.fill.kind === "solid" ? resolveColor(chart.chartBorder.fill.color, theme) : "transparent"}
                    onPick={c => store.updateShapes([chart.id], s => s.kind === "chart"
                      ? { ...s, chartBorder: { fill: { kind: "solid", color: c }, widthPt: s.chartBorder?.widthPt ?? 1, dash: s.chartBorder?.dash } }
                      : s)}
                    onNone={() => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, chartBorder: { fill: { kind: "none" }, widthPt: 1 } } : s)}
                    noneLabel="No Border"
                    theme={theme}
                  />
                  <select
                    className="pane-select"
                    value={String(chart.chartBorder?.widthPt ?? 1)}
                    onChange={e => store.updateShapes([chart.id], s => s.kind === "chart"
                      ? { ...s, chartBorder: { fill: s.chartBorder?.fill ?? { kind: "solid", color: { kind: "srgb", hex: "D9D9D9" } }, widthPt: parseFloat(e.target.value), dash: s.chartBorder?.dash } }
                      : s)}
                  >
                    {[0.75, 1, 1.5, 2, 3].map(w => <option key={w} value={String(w)}>{w} pt</option>)}
                  </select>
                </div>
                <div className="pane-row" style={{ marginTop: 6 }}>
                  <span className="pane-mini-label">Plot fill</span>
                  <InlineColor
                    hex={chart.plotFill?.kind === "solid" ? resolveColor(chart.plotFill.color, theme) : "transparent"}
                    onPick={c => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, plotFill: { kind: "solid", color: c } } : s)}
                    onNone={() => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, plotFill: undefined } : s)}
                    noneLabel="Automatic"
                    theme={theme}
                  />
                  <span className="pane-mini-label">Border</span>
                  <InlineColor
                    hex={chart.plotBorder && chart.plotBorder.fill.kind === "solid" ? resolveColor(chart.plotBorder.fill.color, theme) : "transparent"}
                    onPick={c => store.updateShapes([chart.id], s => s.kind === "chart"
                      ? { ...s, plotBorder: { fill: { kind: "solid", color: c }, widthPt: s.plotBorder?.widthPt ?? 1, dash: s.plotBorder?.dash } }
                      : s)}
                    onNone={() => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, plotBorder: undefined } : s)}
                    noneLabel="No Border"
                    theme={theme}
                  />
                </div>
                {!["pie", "doughnut"].includes(chart.chart) && (
                  <div className="pane-row" style={{ marginTop: 6 }}>
                    <label className="pane-check" style={{ margin: 0 }}>
                      <input
                        type="checkbox"
                        checked={!chart.hideGridlines}
                        onChange={e => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, hideGridlines: e.target.checked ? undefined : true } : s)}
                      />
                      Gridlines
                    </label>
                    <InlineColor
                      hex={chart.gridColor ? resolveColor(chart.gridColor, theme) : "#E0E0E0"}
                      onPick={c => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, gridColor: c, hideGridlines: undefined } : s)}
                      onNone={() => store.updateShapes([chart.id], s => s.kind === "chart" ? { ...s, gridColor: undefined } : s)}
                      noneLabel="Automatic"
                      theme={theme}
                    />
                  </div>
                )}
              </div>
              </Acc>
              {["line", "scatter", "radar"].includes(chart.chart) && (
                <Acc id="chart-series-lines" title="Series lines &amp; markers" defaultOpen>
                <div className="pane-section">
                  <div className="pane-row">
                    <span className="pane-mini-label">Width</span>
                    <select
                      className="pane-select"
                      value={String(chart.series[0]?.lineWidthPt ?? 2.25)}
                      onChange={e => {
                        const v = parseFloat(e.target.value);
                        store.updateShapes([chart.id], s => s.kind === "chart"
                          ? { ...s, series: s.series.map(x => ({ ...x, lineWidthPt: v === 2.25 ? undefined : v })) }
                          : s);
                      }}
                    >
                      {[1, 1.5, 2.25, 3, 4, 6].includes(chart.series[0]?.lineWidthPt ?? 2.25)
                        ? null
                        : <option value={String(chart.series[0]?.lineWidthPt)}>{chart.series[0]?.lineWidthPt} pt</option>}
                      {[1, 1.5, 2.25, 3, 4, 6].map(w => <option key={w} value={String(w)}>{w} pt</option>)}
                    </select>
                    <span className="pane-mini-label">Dash</span>
                    <select
                      className="pane-select"
                      value={chart.series[0]?.dash ?? "solid"}
                      onChange={e => {
                        const v = e.target.value as "solid" | "dash" | "dot";
                        store.updateShapes([chart.id], s => s.kind === "chart"
                          ? { ...s, series: s.series.map(x => ({ ...x, dash: v === "solid" ? undefined : v })) }
                          : s);
                      }}
                    >
                      <option value="solid">Solid</option>
                      <option value="dash">Dashed</option>
                      <option value="dot">Dotted</option>
                    </select>
                  </div>
                  {chart.chart !== "radar" && (
                    <div className="pane-row" style={{ marginTop: 6 }}>
                      <span className="pane-mini-label">Marker size</span>
                      <input
                        type="number" min={2} max={20} className="pane-num"
                        value={chart.markerSizePt ?? 5}
                        onChange={e => {
                          const v = parseInt(e.target.value, 10);
                          store.updateShapes([chart.id], s => s.kind === "chart"
                            ? { ...s, markerSizePt: Number.isFinite(v) && v >= 2 && v <= 20 && v !== 5 ? v : undefined }
                            : s);
                        }}
                      />
                      <span className="pane-val">pt</span>
                    </div>
                  )}
                </div>
                </Acc>
              )}
              <Acc id="chart-colors" title={["pie", "doughnut"].includes(chart.chart) ? "Slice colors" : "Series colors"} defaultOpen>
              <div className="pane-section">
                {["pie", "doughnut"].includes(chart.chart)
                  ? chart.categories.map((cat, i) => (
                    <div className="pane-row" key={i} style={{ marginTop: 4 }}>
                      <InlineColor
                        hex={seriesColor(i, theme, chart.pointColors?.[i] ?? undefined)}
                        onPick={c => store.updateShapes([chart.id], sh => sh.kind === "chart" ? { ...sh, pointColors: setSlice(sh, i, c) } : sh)}
                        onNone={() => store.updateShapes([chart.id], sh => sh.kind === "chart" ? { ...sh, pointColors: setSlice(sh, i, null) } : sh)}
                        noneLabel="Automatic"
                        theme={theme}
                      />
                      <span className="pane-mini-label ellip">{cat || `Slice ${i + 1}`}</span>
                    </div>
                  ))
                  : chart.series.map((s, i) => (
                    <div className="pane-row" key={i} style={{ marginTop: 4 }}>
                      <InlineColor
                        hex={seriesColor(i, theme, s.color)}
                        onPick={c => store.updateShapes([chart.id], sh => sh.kind === "chart"
                          ? { ...sh, series: sh.series.map((x, k) => (k === i ? { ...x, color: c } : x)) }
                          : sh)}
                        onNone={() => store.updateShapes([chart.id], sh => sh.kind === "chart"
                          ? { ...sh, series: sh.series.map((x, k) => (k === i ? { ...x, color: undefined } : x)) }
                          : sh)}
                        noneLabel="Automatic"
                        theme={theme}
                      />
                      <span className="pane-mini-label ellip">{s.name || `Series ${i + 1}`}</span>
                    </div>
                  ))}
              </div>
              </Acc>
              <Acc id="chart-text" title="Text elements (title / axes / legend)" defaultOpen>
              <div className="pane-section">
                <div className="tbl-style-row" style={{ flexWrap: "wrap" }}>
                  {([
                    ["title", "Title", !!chart.title],
                    ["axisTitleX", "X Title", !!chart.axisTitleX && !["pie", "doughnut", "radar"].includes(chart.chart)],
                    ["axisTitleY", "Y Title", !!chart.axisTitleY && !["pie", "doughnut", "radar"].includes(chart.chart)],
                    ["legend", "Legend", chart.legend],
                    ["axisLabels", "Labels", !["pie", "doughnut", "radar"].includes(chart.chart)],
                  ] as const).filter(([, , ok]) => ok).map(([part, label]) => (
                    <button
                      key={part}
                      className={`pane-btn sm ${state.chartPartSel?.id === chart.id && state.chartPartSel.part === part ? "primary" : ""}`}
                      onClick={() => store.setChartPart(chart.id, state.chartPartSel?.part === part && state.chartPartSel.id === chart.id ? null : part)}
                    >{label}</button>
                  ))}
                </div>
                {state.chartPartSel?.id === chart.id ? (
                  <div className="pane-mini-label" style={{ marginTop: 6 }}>
                    Use the <b>Home</b> tab (font, size, color, B/I/U) to format the selected element.
                    <button className="pane-btn sm" style={{ marginLeft: 6 }} onClick={() => store.resetChartPart()}>Reset element</button>
                  </div>
                ) : (
                  <div className="pane-mini-label" style={{ marginTop: 6 }}>Pick an element (or click it on the chart), then format it from the Home tab.</div>
                )}
              </div>
              </Acc>
              <Acc id="chart-data" title="Data &amp; layout" defaultOpen>
              <div className="pane-section">
                <button className="pane-btn" onClick={() => store.setState({ chartEditId: chart.id })}>Edit Data…</button>
              </div>
              <SizePos shape={chart} cm={cm} updateNum={updateNum} />
              </Acc>
            </>
          )}
        </div>
      )}
      <div className="right-rail">
        <RailToggle label="Slide settings" active={open === "slide"} enabled={true} onClick={() => setOpen(open === "slide" ? null : "slide")}>
          <path d="M3 4.5h14v11H3z M3 12l4-3.5 3 2.5 3.5-3 3.5 3" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
        </RailToggle>
        <RailToggle label="Shape settings" active={open === "shape"} enabled={sel.length > 0 && !pic && !table && !chart} onClick={() => setOpen(open === "shape" ? null : "shape")}>
          <path d="M4 4h7v7H4z M9 9.5a5 5 0 1 0 7 7 5 5 0 0 0-7-7z" fill="none" stroke="currentColor" strokeWidth="1.4" />
        </RailToggle>
        <RailToggle label="Image settings" active={open === "image"} enabled={!!pic} onClick={() => setOpen(open === "image" ? null : "image")}>
          <path d="M3 4.5h14v11H3z M6.5 8.2a1.2 1.2 0 1 0 0-2.4 1.2 1.2 0 0 0 0 2.4z M4.5 14.5 9 10l3 3 2-2 2 3" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round" />
        </RailToggle>
        <RailToggle label="Table settings" active={open === "table"} enabled={!!table} onClick={() => setOpen(open === "table" ? null : "table")}>
          <path d="M3 4.5h14v11H3z M3 9h14M3 13h14M8.5 4.5v11M13 4.5v11" fill="none" stroke="currentColor" strokeWidth="1.3" />
        </RailToggle>
        <RailToggle label="Chart settings" active={open === "chart"} enabled={!!chart} onClick={() => setOpen(open === "chart" ? null : "chart")}>
          <path d="M3.5 16.5V3.5M3.5 16.5h13" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
          <path d="M6.5 13.5v-4M10 13.5V6M13.5 13.5V9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
        </RailToggle>
      </div>
    </div>
  );
}

/**
 * Hover flyout that escapes scrollable ancestors: the submenu is position:fixed,
 * measured from the row at hover time (a CSS right:100% flyout gets clipped by
 * the settings pane's overflow). DOM containment keeps hover alive when the
 * pointer moves from the row into the fixed submenu.
 */
function Flyout({ label, disabled, children }: { label: string; disabled?: boolean; children: React.ReactNode }) {
  const [pos, setPos] = useState<{ top: number; right: number } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  return (
    <div
      className="flyout"
      ref={ref}
      onMouseEnter={() => {
        if (disabled || !ref.current) return;
        const r = ref.current.getBoundingClientRect();
        // 4px overlap under the row keeps an unbroken hover path into the submenu
        setPos({ top: Math.max(8, r.top - 4), right: window.innerWidth - r.left - 4 });
      }}
      onMouseLeave={() => setPos(null)}
    >
      <button className="menu-item fly-label" disabled={disabled}>{label} <span className="fly-arrow">▸</span></button>
      {pos && !disabled && (
        <div className="flyout-sub" style={{ top: pos.top, right: pos.right }}>
          {children}
        </div>
      )}
    </div>
  );
}

/** PowerPoint-style "Chart Elements" menu with hover flyouts. */
function ChartElementsMenu({ chart }: { chart: ChartShape }) {
  const up = (fn: (s: ChartShape) => ChartShape) =>
    store.updateShapes([chart.id], s => (s.kind === "chart" ? fn(s) : s));
  const plain = chart.chart === "pie" || chart.chart === "doughnut" || chart.chart === "radar";
  const check = (on: boolean | undefined) => <span className="fly-check">{on ? "✓" : ""}</span>;

  return (
    <Dropdown panelClassName="elements-panel" button={<span className="pane-btn-face">Add Chart Element ▾</span>}>
      {close => (
        <div className="menu">
          <Flyout label="Axes" disabled={plain}>
            <button className="menu-item" onClick={() => up(s => ({ ...s, hideAxisX: s.hideAxisX ? undefined : true }))}>
              {check(!chart.hideAxisX)} Horizontal Axis
            </button>
            <button className="menu-item" onClick={() => up(s => ({ ...s, hideAxisY: s.hideAxisY ? undefined : true }))}>
              {check(!chart.hideAxisY)} Vertical Axis
            </button>
          </Flyout>
          <Flyout label="Axis Titles" disabled={plain}>
            <button className="menu-item" onClick={() => {
              const v = prompt("Horizontal axis title:", chart.axisTitleX ?? "");
              if (v !== null) up(s => ({ ...s, axisTitleX: v.trim() || undefined }));
              close();
            }}>{check(!!chart.axisTitleX)} Horizontal…</button>
            <button className="menu-item" onClick={() => {
              const v = prompt("Vertical axis title:", chart.axisTitleY ?? "");
              if (v !== null) up(s => ({ ...s, axisTitleY: v.trim() || undefined }));
              close();
            }}>{check(!!chart.axisTitleY)} Vertical…</button>
            <div className="menu-div" />
            <button className="menu-item" onClick={() => up(s => ({ ...s, axisTitleX: undefined, axisTitleY: undefined }))}>None</button>
          </Flyout>
          <Flyout label="Chart Title">
            <button className="menu-item" onClick={() => up(s => ({ ...s, title: undefined }))}>{check(!chart.title)} None</button>
            <button className="menu-item" onClick={() => {
              const v = prompt("Chart title:", chart.title ?? "Chart Title");
              if (v !== null) up(s => ({ ...s, title: v.trim() || undefined }));
              close();
            }}>{check(!!chart.title)} Above Chart…</button>
          </Flyout>
          <Flyout label="Data Labels">
            <button className="menu-item" onClick={() => up(s => ({ ...s, dataLabels: undefined }))}>{check(!chart.dataLabels)} None</button>
            <button className="menu-item" onClick={() => up(s => ({ ...s, dataLabels: true }))}>{check(chart.dataLabels)} Show {plain ? "Percentages" : "Values"}</button>
          </Flyout>
          <Flyout label="Error Bars" disabled={plain}>
            <button className="menu-item" onClick={() => up(s => ({ ...s, errorBarsPct: undefined }))}>{check(!chart.errorBarsPct)} None</button>
            <button className="menu-item" onClick={() => up(s => ({ ...s, errorBarsPct: 5 }))}>{check(chart.errorBarsPct === 5)} Percentage 5%</button>
            <button className="menu-item" onClick={() => up(s => ({ ...s, errorBarsPct: 10 }))}>{check(chart.errorBarsPct === 10)} Percentage 10%</button>
          </Flyout>
          <Flyout label="Legend">
            <button className="menu-item" onClick={() => up(s => ({ ...s, legend: false }))}>{check(!chart.legend)} None</button>
            {([["r", "Right"], ["b", "Bottom"], ["t", "Top"], ["l", "Left"]] as const).map(([pos, label]) => (
              <button key={pos} className="menu-item" onClick={() => up(s => ({ ...s, legend: true, legendPos: pos === "r" ? undefined : pos }))}>
                {check(chart.legend && (chart.legendPos ?? "r") === pos)} {label}
              </button>
            ))}
          </Flyout>
        </div>
      )}
    </Dropdown>
  );
}

/** Fill editor: type dropdown + per-type controls. */
function FillEditor({ sp, theme }: { sp: SpShape; theme: ColorTheme }) {
  return (
    <FillControls
      label="Fill"
      fill={sp.fill}
      setFill={f => store.updateShapes([sp.id], s => (s.kind === "sp" ? { ...s, fill: f } : s))}
      setFillLive={f => store.updateShapes([sp.id], s => (s.kind === "sp" ? { ...s, fill: f } : s), { historic: false })}
      theme={theme}
    />
  );
}

/** Fill editor: type dropdown + per-kind controls. Drives shape fills and the slide background. */
function FillControls({ label, fill, setFill, setFillLive, theme, types = FILL_TYPES }: {
  label: string;
  fill: Fill;
  setFill: (f: Fill) => void;
  /** Preview path used during slider drags (no history entry per move). */
  setFillLive?: (f: Fill) => void;
  theme: ColorTheme;
  types?: ReadonlyArray<readonly [string, string]>;
}) {
  const fileRef = useRef<HTMLInputElement>(null);
  const allowNone = types.some(([v]) => v === "none");

  const onType = (t: string) => {
    if (t === fill.kind) return;
    const baseColor = fill.kind === "solid" ? fill.color : { kind: "scheme" as const, slot: "accent1" as const };
    switch (t) {
      case "solid": setFill({ kind: "solid", color: baseColor }); break;
      case "gradient":
        setFill({
          kind: "gradient", angle: 90,
          stops: [
            { pos: 0, color: baseColor },
            { pos: 100, color: { kind: "scheme", slot: "accent1", lumMod: 40, lumOff: 60 } },
          ],
        });
        break;
      case "image": fileRef.current?.click(); break; // switch happens after the file loads
      case "pattern": setFill({ kind: "pattern", prst: "ltDnDiag", fg: baseColor, bg: { kind: "srgb", hex: "FFFFFF" } }); break;
      case "none": setFill({ kind: "none" }); break;
    }
  };

  const grad = fill.kind === "gradient" ? fill : null;

  return (
    <div className="pane-section">
      <div className="pane-label">{label}</div>
      <select className="pane-select wide" value={fill.kind} onChange={e => onType(e.target.value)}>
        {types.map(([v, lbl]) => <option key={v} value={v}>{lbl}</option>)}
      </select>

      {fill.kind === "solid" && (
        <div className="pane-row" style={{ marginTop: 6 }}>
          <span className="pane-mini-label">Color</span>
          <InlineColor
            hex={resolveColor(fill.color, theme)}
            onPick={c => setFill({ kind: "solid", color: c })}
            onNone={() => setFill(allowNone ? { kind: "none" } : { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } })}
            noneLabel={allowNone ? "No Fill" : "White"}
            theme={theme}
          />
        </div>
      )}

      {grad && (
        <>
          <GradientSlider
            stops={grad.stops}
            theme={theme}
            onStops={(stops, commit) => {
              if (!commit) {
                (setFillLive ?? setFill)({ ...grad, stops });
              } else if (setFillLive) {
                // route the final state through the preview path so the whole
                // gesture lands as exactly one undo entry (the pre-drag base)
                setFillLive({ ...grad, stops });
                store.endPreview(true);
              } else {
                setFill({ ...grad, stops });
              }
            }}
          />
          <div className="pane-row" style={{ marginTop: 6 }}>
            <select
              className="pane-select"
              title="Direction"
              value={[0, 45, 90, 135, 180, 270].includes(Math.round(grad.angle)) ? String(Math.round(grad.angle)) : ""}
              onChange={e => { if (e.target.value !== "") setFill({ ...grad, angle: parseInt(e.target.value, 10) }); }}
            >
              <option value="" disabled>Custom</option>
              <option value="0">Linear right →</option>
              <option value="45">Linear down-right ↘</option>
              <option value="90">Linear down ↓</option>
              <option value="135">Linear down-left ↙</option>
              <option value="180">Linear left ←</option>
              <option value="270">Linear up ↑</option>
            </select>
            <span className="pane-mini-label">Angle</span>
            <input
              type="number" className="pane-num"
              value={Math.round(grad.angle)}
              onChange={e => setFill({ ...grad, angle: ((parseInt(e.target.value, 10) || 0) % 360 + 360) % 360 })}
            />
            <span className="pane-val">°</span>
          </div>
        </>
      )}

      {fill.kind === "image" && (
        <div className="pane-row" style={{ marginTop: 6 }}>
          <button className="pane-btn sm" onClick={() => fileRef.current?.click()}>Select Picture…</button>
          <select
            className="pane-select"
            value={fill.tile ? "tile" : "stretch"}
            onChange={e => setFill({ ...fill, tile: e.target.value === "tile" || undefined })}
          >
            <option value="stretch">Stretch</option>
            <option value="tile">Tile</option>
          </select>
        </div>
      )}

      {fill.kind === "pattern" && (
        <>
          <div className="pattern-grid" style={{ marginTop: 6 }}>
            {PATTERN_PRESETS.map(p => (
              <button
                key={p}
                className={`pattern-cell ${fill.prst === p ? "active" : ""}`}
                title={p}
                onClick={() => setFill({ ...fill, prst: p })}
              >
                <PatternPreview prst={p} fg={resolveColor(fill.fg, theme)} bg={resolveColor(fill.bg, theme)} />
              </button>
            ))}
          </div>
          <div className="pane-row" style={{ marginTop: 6 }}>
            <span className="pane-mini-label">Fore</span>
            <InlineColor hex={resolveColor(fill.fg, theme)} onPick={c => setFill({ ...fill, fg: c })} onNone={() => undefined} noneLabel="" theme={theme} />
            <span className="pane-mini-label">Back</span>
            <InlineColor hex={resolveColor(fill.bg, theme)} onPick={c => setFill({ ...fill, bg: c })} onNone={() => undefined} noneLabel="" theme={theme} />
          </div>
        </>
      )}

      <input
        ref={fileRef}
        type="file"
        accept="image/png,image/jpeg,image/gif,image/svg+xml"
        style={{ display: "none" }}
        onChange={async e => {
          const f = e.target.files?.[0];
          e.target.value = "";
          if (!f) return;
          const { mediaId } = await loadImageFile(f);
          setFill({ kind: "image", mediaId, tile: fill.kind === "image" ? fill.tile : undefined });
        }}
      />
    </div>
  );
}

function PatternPreview({ prst, fg, bg }: { prst: string; fg: string; bg: string }) {
  const tiles: Record<string, React.ReactNode> = {
    ltDnDiag: <path d="M0 6L6 0M6 12L12 6" stroke={fg} strokeWidth="1" />,
    ltUpDiag: <path d="M0 6L6 12M6 0l6 6" stroke={fg} strokeWidth="1" />,
    dkDnDiag: <path d="M0 6L6 0M6 12L12 6" stroke={fg} strokeWidth="2.2" />,
    horz: <path d="M0 3h12M0 9h12" stroke={fg} strokeWidth="1.2" />,
    vert: <path d="M3 0v12M9 0v12" stroke={fg} strokeWidth="1.2" />,
    cross: <path d="M0 6h12M6 0v12" stroke={fg} strokeWidth="1" />,
    diagCross: <path d="M0 0l12 12M12 0L0 12" stroke={fg} strokeWidth="1" />,
    smGrid: <path d="M0 3h12M0 9h12M3 0v12M9 0v12" stroke={fg} strokeWidth="0.8" />,
  };
  return (
    <svg width="24" height="18" viewBox="0 0 12 12" preserveAspectRatio="none">
      <rect width="12" height="12" fill={bg} />
      {tiles[prst]}
    </svg>
  );
}

function SizePos({ shape, cm, updateNum }: {
  shape: { x: number; y: number; w: number; h: number; rot: number };
  cm: (emu: number) => number;
  updateNum: (f: "x" | "y" | "w" | "h" | "rot", v: number) => void;
}) {
  return (
    <div className="pane-section">
      <div className="pane-label">Size &amp; position (px)</div>
      <div className="pane-grid">
        <label>W <input type="number" value={cm(shape.w)} onChange={e => updateNum("w", parseFloat(e.target.value))} /></label>
        <label>H <input type="number" value={cm(shape.h)} onChange={e => updateNum("h", parseFloat(e.target.value))} /></label>
        <label>X <input type="number" value={cm(shape.x)} onChange={e => updateNum("x", parseFloat(e.target.value))} /></label>
        <label>Y <input type="number" value={cm(shape.y)} onChange={e => updateNum("y", parseFloat(e.target.value))} /></label>
        <label>⟳ <input type="number" value={Math.round(shape.rot)} onChange={e => updateNum("rot", parseFloat(e.target.value))} /></label>
      </div>
    </div>
  );
}

function InlineColor({ hex, onPick, onNone, noneLabel, theme }: {
  hex: string;
  onPick: (c: any) => void;
  onNone: () => void;
  noneLabel: string;
  theme: any;
}) {
  return (
    <Dropdown
      fixed
      panelClassName="color-panel"
      button={<span className="inline-swatch" style={{ background: hex === "transparent" ? "repeating-conic-gradient(#ddd 0 25%, #fff 0 50%) 0 0/10px 10px" : hex }} />}
    >
      {close => <ColorGrid theme={theme} onPick={onPick} allowNone noneLabel={noneLabel} onNone={onNone} close={close} />}
    </Dropdown>
  );
}

function RailToggle({ children, label, active, enabled, onClick }: {
  children: React.ReactNode; label: string; active: boolean; enabled: boolean; onClick: () => void;
}) {
  return (
    <button className={`rail-btn ${active ? "active" : ""}`} title={label} disabled={!enabled} onClick={onClick}>
      <svg width="20" height="20" viewBox="0 0 20 20">{children}</svg>
    </button>
  );
}
