import React from "react";
import type { ArrowEnd, ArrowKind, ColorTheme, Fill, LineDash, MediaItem, Paragraph, Run, Shape, TextBody } from "../model/types";
import { EMU_PER_PX, resolveColor, resolveFill, resolveFontName } from "../model/defaults";

export const px = (emu: number) => emu / EMU_PER_PX;
export const toEmu = (pxv: number) => Math.round(pxv * EMU_PER_PX);
export const ptToPx = (pt: number) => (pt * 96) / 72;

// dash patterns in multiples of the line width (PowerPoint scales dashes with weight)
const DASH_UNITS: Partial<Record<LineDash, number[]>> = {
  dot: [1, 2], sysDot: [1, 1.5], dash: [4, 3], dashDot: [4, 2, 1, 2],
  lgDash: [8, 3], lgDashDot: [8, 2, 1, 2],
};
export function lineDashArray(dash: LineDash | undefined, widthPx: number): string | undefined {
  const u = dash && DASH_UNITS[dash];
  if (!u) return undefined;
  return u.map(n => Math.max(0.5, n * Math.max(0.6, widthPx))).join(" ");
}

// ---- arrowheads (SVG markers; markerUnits=strokeWidth so they scale with weight) ----
const ARROW_SIZE = { sm: 2.4, med: 3.8, lg: 5.6 } as const;
function arrowShape(kind: ArrowKind, color: string): React.ReactNode {
  switch (kind) {
    case "triangle": return <path d="M0 1 L10 5 L0 9 Z" fill={color} />;
    case "stealth": return <path d="M0 1 L10 5 L0 9 L3.4 5 Z" fill={color} />;
    case "arrow": return <path d="M0.4 1 L10 5 L0.4 9 L3 5 Z" fill={color} />;
    case "diamond": return <path d="M0 5 L5 1 L10 5 L5 9 Z" fill={color} />;
    case "oval": return <circle cx="5" cy="5" r="4.2" fill={color} />;
    default: return null;
  }
}
/** A <marker> for one arrow end (auto-start-reverse so the same def serves start & end). */
export function arrowMarker(id: string, end: ArrowEnd | undefined, color: string): React.ReactNode {
  if (!end || end.type === "none") return null;
  const len = ARROW_SIZE[end.len ?? "med"], w = ARROW_SIZE[end.w ?? "med"];
  return (
    <marker
      id={id} viewBox="0 0 10 10" refX={9} refY={5}
      markerWidth={len} markerHeight={w} orient="auto-start-reverse" markerUnits="strokeWidth"
    >
      {arrowShape(end.type, color)}
    </marker>
  );
}

export function shapeTransform(s: Shape): string {
  const x = px(s.x), y = px(s.y), w = px(s.w), h = px(s.h);
  const parts = [`translate(${x} ${y})`];
  if (s.rot) parts.push(`rotate(${s.rot} ${w / 2} ${h / 2})`);
  if (s.flipH) parts.push(`translate(${w} 0) scale(-1 1)`);
  if (s.flipV) parts.push(`translate(0 ${h}) scale(1 -1)`);
  return parts.join(" ");
}

/** SVG <linearGradient> for an OOXML gradient fill (angle: 0° = left→right, cw, y-down). */
export function GradDef({ fill, id, theme }: { fill: Extract<Fill, { kind: "gradient" }>; id: string; theme: ColorTheme }) {
  const rad = (fill.angle * Math.PI) / 180;
  const dx = Math.cos(rad) / 2, dy = Math.sin(rad) / 2;
  return (
    <linearGradient id={id} x1={0.5 - dx} y1={0.5 - dy} x2={0.5 + dx} y2={0.5 + dy}>
      {fill.stops.map((s, i) => (
        <stop key={i} offset={`${s.pos}%`} stopColor={resolveColor(s.color, theme)} />
      ))}
    </linearGradient>
  );
}

/** Hatch path generators for a:pattFill presets (8px tile), approximating PowerPoint. */
const PATTERN_TILES: Record<string, React.ReactNode> = {
  ltDnDiag: <path d="M0 8L8 0" strokeWidth="1" />,
  ltUpDiag: <path d="M0 0L8 8" strokeWidth="1" />,
  dkDnDiag: <path d="M0 8L8 0M-2 2L2 -2M6 10l4-4" strokeWidth="2" />,
  dkUpDiag: <path d="M0 0L8 8M6 -2l4 4M-2 6l4 4" strokeWidth="2" />,
  horz: <path d="M0 4h8" strokeWidth="1.2" />,
  vert: <path d="M4 0v8" strokeWidth="1.2" />,
  cross: <path d="M0 4h8M4 0v8" strokeWidth="1" />,
  diagCross: <path d="M0 0l8 8M8 0L0 8" strokeWidth="1" />,
  smGrid: <path d="M0 2h8M0 6h8M2 0v8M6 0v8" strokeWidth="0.7" />,
  pct20: <><circle cx="2" cy="2" r="1" /><circle cx="6" cy="6" r="1" /></>,
  pct50: <><rect x="0" y="0" width="4" height="4" /><rect x="4" y="4" width="4" height="4" /></>,
};

function PatternDef({ fill, id, theme }: { fill: Extract<Fill, { kind: "pattern" }>; id: string; theme: ColorTheme }) {
  const fg = resolveColor(fill.fg, theme);
  const tile = PATTERN_TILES[fill.prst] ?? PATTERN_TILES.ltDnDiag;
  return (
    <pattern id={id} width="8" height="8" patternUnits="userSpaceOnUse">
      <rect width="8" height="8" fill={resolveColor(fill.bg, theme)} />
      <g stroke={fg} fill={fg}>{tile}</g>
    </pattern>
  );
}

function ImageFillDef({ fill, id, media }: { fill: Extract<Fill, { kind: "image" }>; id: string; media?: Map<string, MediaItem> }) {
  const item = media?.get(fill.mediaId);
  if (!item) return <pattern id={id} width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="#D9D9D9" /></pattern>;
  if (fill.tile) {
    return (
      <pattern id={id} width="96" height="96" patternUnits="userSpaceOnUse">
        <image href={item.dataUrl} width="96" height="96" preserveAspectRatio="xMidYMid slice" />
      </pattern>
    );
  }
  // stretch to the shape's bounding box
  return (
    <pattern id={id} width="1" height="1" patternContentUnits="objectBoundingBox">
      <image href={item.dataUrl} width="1" height="1" preserveAspectRatio="none" />
    </pattern>
  );
}

/** Returns [paint, defs] — gradient/image/pattern fills get a generated <defs> entry. */
export function paintFor(fill: Fill | undefined, theme: ColorTheme, id: string, media?: Map<string, MediaItem>): [string, React.ReactNode] {
  if (fill?.kind === "gradient") {
    return [`url(#${id})`, <GradDef key={id} fill={fill} id={id} theme={theme} />];
  }
  if (fill?.kind === "pattern") {
    return [`url(#${id})`, <PatternDef key={id} fill={fill} id={id} theme={theme} />];
  }
  if (fill?.kind === "image") {
    return [`url(#${id})`, <ImageFillDef key={id} fill={fill} id={id} media={media} />];
  }
  return [resolveFill(fill, theme), null];
}

export function runStyle(r: Run, theme: ColorTheme): React.CSSProperties {
  const family = resolveFontName(r.font, theme);
  // graceful substitution when the deck's font isn't installed (PowerPoint substitutes too)
  const narrow = /condensed|narrow/i.test(family) ? "'Arial Narrow', " : "";
  return {
    fontFamily: `'${family}', ${narrow}'Segoe UI', Calibri, Arial, sans-serif`,
    fontSize: `${ptToPx(r.sizePt)}px`,
    fontWeight: r.bold ? 700 : 400,
    fontStyle: r.italic ? "italic" : "normal",
    textDecoration: [r.underline ? "underline" : "", r.strike ? "line-through" : ""].filter(Boolean).join(" ") || "none",
    color: resolveColor(r.color, theme),
    backgroundColor: r.highlight ? resolveColor(r.highlight, theme) : undefined,
    verticalAlign: r.baseline ? `${r.baseline}%` : undefined,
    // a:rPr@cap — display only (the underlying text keeps its case)
    textTransform: r.caps === "all" ? "uppercase" : undefined,
    fontVariant: r.caps === "small" ? "small-caps" : undefined,
    whiteSpace: "pre-wrap",
  };
}

const ALIGN_MAP = { l: "left", ctr: "center", r: "right", just: "justify" } as const;

export function paraStyle(p: Paragraph): React.CSSProperties {
  const indent = p.level * 28 + (p.bullet !== "none" ? 24 : 0);
  return {
    textAlign: ALIGN_MAP[p.align],
    // OOXML spcPct is relative to SINGLE spacing (~1.2em), not to the font size —
    // raw CSS percentages render ~25% too tight (the "cramped paragraph" bug)
    lineHeight: p.lineSpacingPct ? (p.lineSpacingPct / 100) * 1.2 : 1.2,
    paddingLeft: indent ? `${indent}px` : undefined,
    textIndent: p.bullet !== "none" ? "-24px" : undefined,
    margin: 0,
    minHeight: "1em",
  };
}

export function bulletPrefix(p: Paragraph, numIndex: number): string {
  if (p.bullet === "char") return "•  ";
  if (p.bullet === "num") return `${numIndex}.  `;
  return "";
}

function paraFallbackSize(p: Paragraph, body: TextBody): number {
  if (p.runs.length) return p.runs[0].sizePt;
  const first = body.paragraphs.find(q => q.runs.length);
  return first?.runs[0]?.sizePt ?? 18;
}

/** The laid-out text block (HTML), shared by shape text bodies and table cells. */
export function TextContent({ body, w, h, theme }: { body: TextBody; w: number; h: number; theme: ColorTheme }) {
  const [il, it, ir, ib] = body.insets.map(px);
  const justify = body.anchor === "t" ? "flex-start" : body.anchor === "b" ? "flex-end" : "center";
  let numIndex = 0;
  const cols = body.columns && body.columns > 1 ? body.columns : 0;
  const paras = body.paragraphs.map((p, i) => {
    numIndex = p.bullet === "num" ? numIndex + 1 : 0;
    const prefix = bulletPrefix(p, numIndex);
    return (
      <div key={i} style={paraStyle(p)}>
        {p.runs.length === 0 || p.runs.every(r => !r.text) ? (
          <span style={p.runs.length ? runStyle(p.runs[0], theme) : { fontSize: `${ptToPx(paraFallbackSize(p, body))}px` }}>{"​"}</span>
        ) : (
          <>
            {prefix && <span style={{ ...runStyle(p.runs[0], theme), textDecoration: "none" }}>{prefix}</span>}
            {p.runs.map((r, j) => (
              <span key={j} style={runStyle(r, theme)}>{r.text}</span>
            ))}
          </>
        )}
      </div>
    );
  });
  return (
    <div
      style={{
        width: `${w}px`, height: `${h}px`, boxSizing: "border-box",
        padding: `${it}px ${ir}px ${ib}px ${il}px`,
        display: "flex", flexDirection: "column", justifyContent: justify,
        overflow: "visible", wordBreak: "break-word",
      }}
    >
      {cols ? (
        <div style={{ columnCount: cols, columnGap: `${px(body.colSpacing ?? 360000)}px`, columnFill: "auto", maxHeight: "100%" }}>
          {paras}
        </div>
      ) : paras}
    </div>
  );
}
