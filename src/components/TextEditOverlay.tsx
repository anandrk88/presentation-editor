import React, { useEffect, useLayoutEffect, useRef } from "react";
import type { ColorRef, ColorTheme, Paragraph, Run, SpShape, TableShape, TextBody } from "../model/types";
import { resolveColor, resolveFontName } from "../model/defaults";
import { bulletPrefix, paraStyle, px, runStyle } from "../render/SlideView";
import { tableGrid } from "../render/GraphicViews";
import { store } from "../state/store";
import { editorBus } from "../state/useStore";

/**
 * In-place rich text editor: a contentEditable div positioned exactly over the
 * shape (or table cell), in slide-pixel coordinates scaled by the current zoom.
 * On commit the DOM is parsed back into paragraphs/runs (computed styles -> run props).
 */
export function TextEditOverlay({ shape, cell, zoom, theme }: {
  shape: SpShape | TableShape;
  cell?: { row: number; col: number } | null;
  zoom: number;
  theme: ColorTheme;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const committed = useRef(false);

  // resolve editing target geometry + body
  let left = px(shape.x), top = px(shape.y), w = px(shape.w), h = px(shape.h);
  let body: TextBody | undefined;
  if (shape.kind === "table") {
    const c = cell ?? { row: 0, col: 0 };
    const grid = tableGrid(shape);
    const tcell = shape.cells[c.row]?.[c.col];
    const span = Math.max(1, tcell?.gridSpan ?? 1);
    const vspan = Math.max(1, tcell?.rowSpan ?? 1);
    left += grid.colX[c.col];
    top += grid.rowY[c.row];
    w = grid.colW.slice(c.col, c.col + span).reduce((a, b) => a + b, 0);
    h = grid.rowH.slice(c.row, c.row + vspan).reduce((a, b) => a + b, 0);
    body = tcell?.text;
  } else {
    body = shape.text;
  }

  const [il, it, ir, ib] = (body?.insets ?? [91440, 45720, 91440, 45720]).map(px);
  const anchor = body?.anchor ?? "t";
  const justify = anchor === "t" ? "flex-start" : anchor === "b" ? "flex-end" : "center";

  // style seed: without it, text typed into empty paragraphs computes the app's
  // 11px UI font and commits as ~8.5pt — the "tiny invisible text" bug
  const seedRun: Run = body?.paragraphs.flatMap(p => p.runs)[0]
    ?? { text: "", sizePt: 18, font: "+mn-lt", color: { kind: "scheme", slot: "dk1" } };

  const commit = () => {
    if (committed.current || !ref.current) return;
    committed.current = true;
    const paragraphs = parseEditedDom(ref.current, body, theme, seedRun);
    if (shape.kind === "table") {
      const c = cell ?? { row: 0, col: 0 };
      store.updateShapes([shape.id], s => {
        if (s.kind !== "table") return s;
        const cells = s.cells.map((row, ri) => ri !== c.row ? row : row.map((tc, ci) =>
          ci !== c.col ? tc : { ...tc, text: { ...tc.text, paragraphs } }));
        return { ...s, cells };
      });
    } else {
      store.updateShapes([shape.id], s =>
        s.kind === "sp" ? { ...s, text: { ...(s.text ?? { anchor: "t", wrap: true, insets: [91440, 45720, 91440, 45720] }), paragraphs } } : s
      );
    }
    store.setState({ editingShapeId: null, editingCell: null });
  };

  useEffect(() => {
    editorBus.commitTextEdit = commit;
    return () => { editorBus.commitTextEdit = null; };
  });

  useLayoutEffect(() => {
    document.execCommand("styleWithCSS", false, "true");
    const el = ref.current;
    if (!el) return;
    el.focus();
    const sel = window.getSelection();
    if (sel) {
      const range = document.createRange();
      range.selectNodeContents(el);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    }
  }, []);

  const paras = body?.paragraphs ?? [];
  let numIndex = 0;

  return (
    <div
      className="text-edit-wrap"
      style={{
        position: "absolute",
        left: left * zoom,
        top: top * zoom,
        width: w * zoom,
        height: h * zoom,
        transform: shape.kind !== "table" && shape.rot ? `rotate(${shape.rot}deg)` : undefined,
        zIndex: 5,
      }}
      onPointerDown={e => e.stopPropagation()}
    >
      <div
        ref={ref}
        className="text-edit"
        contentEditable
        suppressContentEditableWarning
        spellCheck={false}
        style={{
          width: w, height: h,
          transform: `scale(${zoom})`, transformOrigin: "0 0",
          boxSizing: "border-box",
          padding: `${it}px ${ir}px ${ib}px ${il}px`,
          display: "flex", flexDirection: "column", justifyContent: justify,
          outline: "none", overflow: "visible", wordBreak: "break-word", cursor: "text",
          ...runStyle(seedRun, theme), // container-level seed so stray text nodes inherit sanely
        }}
        onBlur={commit}
        onKeyDown={e => {
          e.stopPropagation();
          if (e.key === "Escape") { e.preventDefault(); commit(); }
        }}
      >
        {wrapColumns(body, paras.map((p, i) => {
          numIndex = p.bullet === "num" ? numIndex + 1 : 0;
          const prefix = bulletPrefix(p, numIndex);
          // first-run style on the block keeps stray text nodes (typed at boundaries) styled correctly
          const blockStyle = { ...runStyle(p.runs[0] ?? seedRun, theme), ...paraStyle(p) };
          const visibleRuns = p.runs.filter(r => r.text.length > 0);
          return (
            <div key={i} data-para={i} data-prefix={prefix} className={prefix ? "with-bullet" : undefined} style={blockStyle}>
              {visibleRuns.length === 0
                ? <br />
                : visibleRuns.map((r, j) => (
                  // keyword super/sub (not %) so computed styles read back cleanly on commit
                  <span
                    key={j}
                    data-run="1"
                    style={{ ...runStyle(r, theme), verticalAlign: r.baseline ? (r.baseline > 0 ? "super" : "sub") : undefined }}
                  >{r.text}</span>
                ))}
            </div>
          );
        }))}
      </div>
    </div>
  );
}

/** Match TextContent's column wrapper so commit can find the paragraph container. */
function wrapColumns(body: TextBody | undefined, paras: React.ReactNode): React.ReactNode {
  const cols = body?.columns && body.columns > 1 ? body.columns : 0;
  if (!cols) return paras;
  return (
    <div style={{ columnCount: cols, columnGap: `${px(body!.colSpacing ?? 360000)}px`, columnFill: "auto", maxHeight: "100%" }}>
      {paras}
    </div>
  );
}

// ---------- DOM -> model ----------
function rgbToHex(rgb: string): string {
  const m = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!m) return "000000";
  return [m[1], m[2], m[3]].map(v => parseInt(v, 10).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function styleKey(font: string, sizePt: number, bold: boolean, italic: boolean, underline: boolean, strike: boolean, hex: string) {
  return `${font}|${sizePt}|${bold ? 1 : 0}${italic ? 1 : 0}${underline ? 1 : 0}${strike ? 1 : 0}|${hex}`;
}

function baselineOf(cs: CSSStyleDeclaration): number | undefined {
  const va = cs.verticalAlign;
  if (va === "super") return 30;
  if (va === "sub") return -25;
  if (va.endsWith("%")) {
    const v = parseFloat(va);
    return v ? v : undefined;
  }
  if (va.endsWith("px")) {
    const pxv = parseFloat(va);
    if (Math.abs(pxv) < 0.5) return undefined;
    return pxv > 0 ? 30 : -25; // resolved % — snap to the standard offsets
  }
  return undefined;
}

function highlightOf(cs: CSSStyleDeclaration): { kind: "srgb"; hex: string } | undefined {
  const bg = cs.backgroundColor;
  if (!bg || bg === "transparent" || /rgba?\(\s*0,\s*0,\s*0,\s*0\s*\)/.test(bg)) return undefined;
  return { kind: "srgb", hex: rgbToHex(bg) };
}

function parseEditedDom(root: HTMLElement, body: TextBody | undefined, theme: ColorTheme, seedRun?: Run): Paragraph[] {
  const origParas = body?.paragraphs ?? [];
  // index original runs (keyed by their RESOLVED style) so theme color refs and
  // +mj-lt/+mn-lt font refs survive the computed-style round-trip unchanged
  const origRunByKey = new Map<string, { color: ColorRef; font: string }>();
  const indexRun = (r: Run) => {
    origRunByKey.set(
      styleKey(resolveFontName(r.font, theme).toLowerCase(), r.sizePt, !!r.bold, !!r.italic, !!r.underline, !!r.strike, hexOf(r.color, theme)),
      { color: r.color, font: r.font },
    );
  };
  for (const p of origParas) for (const r of p.runs) indexRun(r);
  if (seedRun) indexRun(seedRun);

  const fallbackRun: Run = origParas.flatMap(p => p.runs)[0] ?? { text: "", sizePt: 18, font: "Arial", color: { kind: "scheme", slot: "dk1" } };

  // paragraphs may sit inside a column wrapper — find their real container
  const container = (root.querySelector("[data-para]")?.parentElement as HTMLElement | null) ?? root;
  const blocks: HTMLElement[] = [];
  let hasBlocks = false;
  container.childNodes.forEach(n => {
    if (n.nodeType === 1 && /^(DIV|P)$/i.test((n as Element).tagName)) { blocks.push(n as HTMLElement); hasBlocks = true; }
  });
  if (!hasBlocks) blocks.push(container); // everything inline — treat as a single paragraph

  const paras: Paragraph[] = [];
  blocks.forEach((block, bi) => {
    const orig = origParas[Math.min(bi, Math.max(0, origParas.length - 1))];
    const runs: Run[] = [];
    collectRuns(block, runs);
    const merged: Run[] = [];
    for (const r of runs) {
      const last = merged[merged.length - 1];
      if (last && sameStyle(last, r)) last.text += r.text;
      else merged.push({ ...r });
    }
    const cleaned = merged
      .map(r => ({ ...r, text: r.text.replace(/​/g, "") }))
      .filter(r => r.text.length > 0);
    // keep one style-carrier run when the block went empty, so the style sticks
    if (!cleaned.length && orig?.runs.length) {
      cleaned.push({ ...orig.runs[0], text: "" });
    }
    paras.push({
      runs: cleaned,
      align: orig?.align ?? "l",
      bullet: orig?.bullet ?? "none",
      level: orig?.level ?? 0,
      lineSpacingPct: orig?.lineSpacingPct,
    });
  });

  if (!paras.length) paras.push({ runs: [], align: "l", bullet: "none", level: 0 });
  return paras;

  function collectRuns(el: Node, out: Run[]) {
    el.childNodes.forEach(n => {
      if (n.nodeType === 3) {
        const text = n.textContent ?? "";
        if (!text) return;
        const parent = (n.parentElement ?? root) as HTMLElement;
        const cs = window.getComputedStyle(parent);
        const sizePt = Math.round((parseFloat(cs.fontSize) * 0.75) * 2) / 2;
        const font = (cs.fontFamily.split(",")[0] ?? "Arial").replace(/['"]/g, "").trim() || "Arial";
        const bold = parseInt(cs.fontWeight, 10) >= 600;
        const italic = cs.fontStyle === "italic";
        const underline = cs.textDecorationLine.includes("underline");
        const strike = cs.textDecorationLine.includes("line-through");
        const hex = rgbToHex(cs.color);
        const key = styleKey(font.toLowerCase(), sizePt, bold, italic, underline, strike, hex);
        const orig = origRunByKey.get(key);
        const color: ColorRef = orig?.color ?? { kind: "srgb", hex };
        out.push({
          text,
          font: orig?.font ?? font,
          sizePt: sizePt || fallbackRun.sizePt,
          bold: bold || undefined,
          italic: italic || undefined,
          underline: underline || undefined,
          strike: strike || undefined,
          baseline: baselineOf(cs),
          highlight: highlightOf(cs),
          color,
        });
      } else if (n.nodeType === 1) {
        const tag = (n as Element).tagName;
        if (tag === "BR") {
          if (out.length) out[out.length - 1].text += "\n";
          return;
        }
        collectRuns(n, out);
      }
    });
  }
}

function hexOf(c: ColorRef, theme: ColorTheme): string {
  const v = resolveColor(c, theme);
  if (v.startsWith("#")) return v.slice(1).toUpperCase();
  const m = v.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!m) return "000000";
  return [m[1], m[2], m[3]].map(x => parseInt(x, 10).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function sameStyle(a: Run, b: Run): boolean {
  return a.font === b.font && a.sizePt === b.sizePt && !!a.bold === !!b.bold && !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline && !!a.strike === !!b.strike &&
    (a.baseline ?? 0) === (b.baseline ?? 0) &&
    JSON.stringify(a.highlight ?? null) === JSON.stringify(b.highlight ?? null) &&
    JSON.stringify(a.color) === JSON.stringify(b.color);
}
