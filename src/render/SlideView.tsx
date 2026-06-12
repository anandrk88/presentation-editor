import React from "react";
import type {
  ColorTheme, MediaItem, Paragraph, PicShape, Presentation, Shape,
  SlideModel, SpShape,
} from "../model/types";
import { resolveColor, resolveFill } from "../model/defaults";
import { cssColorToHex, tintSvgText } from "../util/svgTint";
import { isLineGeom } from "./geometry";
import { presetOutline, presetPaths } from "./presetGeom";
import { paintFor, paraStyle, ptToPx, px, shapeTransform, TextContent } from "./base";
import { ChartView, TableView } from "./GraphicViews";

// shared render helpers re-exported for canvas/overlay/panel code
export * from "./base";

function promptFor(shape: SpShape): string | null {
  if (!shape.isTextBox) return null;
  const n = shape.name.toLowerCase();
  if (n.includes("subtitle")) return "Click to add subtitle";
  if (n.includes("title")) return "Click to add title";
  if (n.includes("placeholder")) return "Click to add text";
  return null;
}

export function TextBodyView({ shape, theme, showPrompts }: { shape: SpShape; theme: ColorTheme; showPrompts?: boolean }) {
  const body = shape.text;
  if (!body) return null;
  const w = px(shape.w), h = px(shape.h);
  const [il, it, ir, ib] = body.insets.map(px);
  const justify = body.anchor === "t" ? "flex-start" : body.anchor === "b" ? "flex-end" : "center";

  const isEmpty = body.paragraphs.every(p => p.runs.length === 0 || p.runs.every(r => !r.text));
  const prompt = showPrompts && isEmpty ? promptFor(shape) : null;
  if (prompt) {
    const p0 = body.paragraphs[0];
    const size = Math.min(p0?.runs[0]?.sizePt ?? (shape.name.toLowerCase().includes("title") ? 36 : 20), 36);
    return (
      <foreignObject x={0} y={0} width={Math.max(w, 1)} height={Math.max(h, 1)} style={{ overflow: "visible", pointerEvents: "none" }}>
        <div style={{
          width: `${w}px`, height: `${h}px`, boxSizing: "border-box",
          padding: `${it}px ${ir}px ${ib}px ${il}px`,
          display: "flex", flexDirection: "column", justifyContent: justify,
        }}>
          <div style={{ ...paraStyle(p0 ?? { runs: [], align: "l", bullet: "none", level: 0 }), color: "#A6A6A6", fontFamily: "'Calibri Light', Calibri, Arial, sans-serif", fontSize: `${ptToPx(size)}px` }}>
            {prompt}
          </div>
        </div>
      </foreignObject>
    );
  }
  return (
    <foreignObject x={0} y={0} width={Math.max(w, 1)} height={Math.max(h, 1)} style={{ overflow: "visible", pointerEvents: "none" }}>
      <TextContent body={body} w={w} h={h} theme={theme} />
    </foreignObject>
  );
}

function SpView({ shape, theme, media, hideText, showPrompts, defsPrefix }: { shape: SpShape; theme: ColorTheme; media?: Map<string, MediaItem>; hideText?: boolean; showPrompts?: boolean; defsPrefix: string }) {
  const w = px(shape.w), h = px(shape.h);
  const [fill, fillDef] = paintFor(shape.fill, theme, `${defsPrefix}-f-${shape.id}`, media);
  const stroke = resolveFill(shape.line.fill, theme);
  const lineOnly = isLineGeom(shape.geom) && !shape.custPath;
  const strokeProps = {
    stroke,
    strokeWidth: stroke === "none" ? 0 : ptToPx(shape.line.widthPt),
    strokeDasharray: shape.line.dash === "dash" ? "6 4" : shape.line.dash === "dot" ? "2 3" : undefined,
  };
  return (
    <g transform={shapeTransform(shape)}>
      {fillDef && <defs>{fillDef}</defs>}
      {shape.custPath ? (
        <path
          d={shape.custPath.d}
          transform={`scale(${(Math.max(w, 0.01) / shape.custPath.w) || 1} ${(Math.max(h, 0.01) / shape.custPath.h) || 1})`}
          fill={fill}
          vectorEffect="non-scaling-stroke"
          {...strokeProps}
        />
      ) : (
        (presetPaths(shape.geom, Math.max(w, 0.5), Math.max(h, 0.5), shape.adj) ?? [{ d: `M0 0H${w}V${h}H0Z`, noFill: false, noStroke: false }]).map((p, i) => (
          <path
            key={i}
            d={p.d}
            fillRule="evenodd"
            fill={lineOnly || p.noFill ? "none" : fill}
            opacity={p.darken ? 0.85 : p.lighten ? 0.95 : undefined}
            strokeLinecap={lineOnly ? "round" : undefined}
            {...strokeProps}
            stroke={p.noStroke ? "none" : strokeProps.stroke}
          />
        ))
      )}
      {!lineOnly && !hideText && <TextBodyView shape={shape} theme={theme} showPrompts={showPrompts} />}
    </g>
  );
}

/** Recolored svg data-urls, keyed mediaId|hex (tint applied at render time). */
const tintCache = new Map<string, string>();
function tintedSvgUrl(m: MediaItem, hex: string): string {
  const key = `${m.id}|${hex}`;
  let url = tintCache.get(key);
  if (!url) {
    const text = new TextDecoder().decode(m.bytes);
    url = "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(tintSvgText(text, hex))));
    tintCache.set(key, url);
  }
  return url;
}

function PicView({ shape, media, theme, defsPrefix }: { shape: PicShape; media: Map<string, MediaItem>; theme: ColorTheme; defsPrefix: string }) {
  const w = px(shape.w), h = px(shape.h);
  const m0 = media.get(shape.mediaId);
  const m = m0 && m0.mime === "image/svg+xml" && shape.svgTint
    ? { ...m0, dataUrl: tintedSvgUrl(m0, cssColorToHex(resolveColor(shape.svgTint, theme))) }
    : m0;
  const sr = shape.srcRect;
  let img: React.ReactNode = null;
  if (m) {
    if (sr) {
      // crop: scale the full image up so the kept window exactly fills the frame
      const kw = Math.max(0.02, 1 - sr.l - sr.r);
      const kh = Math.max(0.02, 1 - sr.t - sr.b);
      const fw = w / kw, fh = h / kh;
      img = <image href={m.dataUrl} x={-sr.l * fw} y={-sr.t * fh} width={fw} height={fh} preserveAspectRatio="none" />;
    } else {
      img = <image href={m.dataUrl} x={0} y={0} width={w} height={h} preserveAspectRatio="none" />;
    }
  }
  const frameGeom = shape.geom && shape.geom !== "rect" ? shape.geom : null;
  const needClip = !!sr || !!frameGeom;
  const clipId = `${defsPrefix}-clip-${shape.id}`;
  const clipD = frameGeom
    ? presetOutline(frameGeom, Math.max(w, 0.5), Math.max(h, 0.5), shape.adj)
    : `M0 0H${w}V${h}H0Z`;
  return (
    <g transform={shapeTransform(shape)}>
      {needClip && <defs><clipPath id={clipId}><path d={clipD} /></clipPath></defs>}
      {m ? (
        <g clipPath={needClip ? `url(#${clipId})` : undefined}>{img}</g>
      ) : (
        <rect width={w} height={h} fill="#ddd" stroke="#999" />
      )}
    </g>
  );
}

export function ShapeView({ shape, theme, media, hideText, hideCell, showPrompts, defsPrefix = "g" }: {
  shape: Shape;
  theme: ColorTheme;
  media: Map<string, MediaItem>;
  hideText?: boolean;
  hideCell?: { row: number; col: number } | null;
  showPrompts?: boolean;
  defsPrefix?: string;
}) {
  switch (shape.kind) {
    case "pic": return <PicView shape={shape} media={media} theme={theme} defsPrefix={defsPrefix} />;
    case "table": return <TableView shape={shape} theme={theme} hideCell={hideText ? hideCell : null} />;
    case "chart": return <ChartView shape={shape} theme={theme} defsPrefix={defsPrefix} />;
    default: return <SpView shape={shape} theme={theme} media={media} hideText={hideText} showPrompts={showPrompts} defsPrefix={defsPrefix} />;
  }
}

export interface SlideSVGProps {
  pres: Presentation;
  slide: SlideModel;
  media: Map<string, MediaItem>;
  width?: number | string;
  height?: number | string;
  className?: string;
  hiddenShapeId?: string | null;            // shape whose text is being edited in the overlay
  hiddenCell?: { row: number; col: number } | null; // cell being edited (table editing)
  showPrompts?: boolean;                    // editor-only "Click to add title" ghosts
  children?: React.ReactNode;               // extra layers (selection chrome)
  svgRef?: React.Ref<SVGSVGElement>;
  onPointerDown?: (e: React.PointerEvent<SVGSVGElement>) => void;
}

export function SlideSVG({ pres, slide, media, width, height, className, hiddenShapeId, hiddenCell, showPrompts, children, svgRef, onPointerDown }: SlideSVGProps) {
  const uid = React.useId().replace(/[^a-zA-Z0-9_-]/g, "");
  const W = px(pres.slideWidth), H = px(pres.slideHeight);
  const [bgPaint, bgDef] = paintFor(slide.background, pres.theme, `${uid}-bg`, media);
  return (
    <svg
      ref={svgRef}
      viewBox={`0 0 ${W} ${H}`}
      width={width}
      height={height}
      className={className}
      onPointerDown={onPointerDown}
      style={{ display: "block" }}
    >
      {bgDef && <defs>{bgDef}</defs>}
      <rect x={0} y={0} width={W} height={H} fill={bgPaint === "none" ? "#FFFFFF" : bgPaint} data-slide-bg="1" />
      {slide.shapes.map(s => (
        <g key={s.id} data-shape-id={s.id}>
          <ShapeView
            shape={s}
            theme={pres.theme}
            media={media}
            hideText={s.id === hiddenShapeId}
            hideCell={hiddenCell}
            showPrompts={showPrompts}
            defsPrefix={uid}
          />
        </g>
      ))}
      {children}
    </svg>
  );
}
