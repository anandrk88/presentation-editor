import React from "react";
import type { ChartPart, ChartShape, ColorTheme, SchemeSlot, TableShape } from "../model/types";
import { resolveColor, resolveFill, resolveFontName, tableBorderColor, tableCellStyle } from "../model/defaults";
import { paintFor, ptToPx, px, shapeTransform, TextContent } from "./base";

const ACCENTS: SchemeSlot[] = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

export function seriesColor(i: number, theme: ColorTheme, explicit?: Parameters<typeof resolveColor>[0]): string {
  return resolveColor(explicit ?? { kind: "scheme", slot: ACCENTS[i % ACCENTS.length] }, theme);
}

/** Geometry of a table in shape-local px space (weights normalized to w/h). */
export function tableGrid(shape: TableShape) {
  const w = px(shape.w), h = px(shape.h);
  const sumW = shape.colW.reduce((a, b) => a + b, 0) || 1;
  const sumH = shape.rowH.reduce((a, b) => a + b, 0) || 1;
  const colW = shape.colW.map(c => (c / sumW) * w);
  const rowH = shape.rowH.map(r => (r / sumH) * h);
  const colX: number[] = [0];
  for (const c of colW) colX.push(colX[colX.length - 1] + c);
  const rowY: number[] = [0];
  for (const r of rowH) rowY.push(rowY[rowY.length - 1] + r);
  return { w, h, colW, rowH, colX, rowY };
}

/** Style-derived cell fill for the active style family/accent. */
export function cellStyleFill(shape: TableShape, r: number, c: number, theme: ColorTheme): string {
  const f = tableCellStyle(shape, r, c).fill;
  return resolveFill(f, theme) === "none" ? "#FFFFFF" : resolveFill(f, theme);
}

export function TableView({ shape, theme, hideCell }: { shape: TableShape; theme: ColorTheme; hideCell?: { row: number; col: number } | null }) {
  const { w, h, colW, rowH, colX, rowY } = tableGrid(shape);
  const borderMode = shape.borderMode ?? "all";
  const border = resolveColor(tableBorderColor(shape), theme);
  return (
    <g transform={shapeTransform(shape)}>
      {shape.cells.map((row, r) =>
        row.map((cell, c) => {
          if (cell.merged) return null;
          const span = Math.max(1, cell.gridSpan ?? 1);
          const vspan = Math.max(1, cell.rowSpan ?? 1);
          const cw = colW.slice(c, c + span).reduce((a, b) => a + b, 0);
          const ch = rowH.slice(r, r + vspan).reduce((a, b) => a + b, 0);
          const fill = cell.fill ? resolveColor(cell.fill.kind === "solid" ? cell.fill.color : { kind: "srgb", hex: "FFFFFF" }, theme) : cellStyleFill(shape, r, c, theme);
          const hidden = hideCell && hideCell.row === r && hideCell.col === c;
          return (
            <g key={`${r}-${c}`} transform={`translate(${colX[c]} ${rowY[r]})`}>
              <rect
                width={cw} height={ch}
                fill={cell.fill?.kind === "none" ? "none" : fill}
                stroke={borderMode === "all" ? border : "none"}
                strokeWidth={borderMode === "all" ? 1.4 : 0}
              />
              {!hidden && (
                <foreignObject x={0} y={0} width={Math.max(cw, 1)} height={Math.max(ch, 1)} style={{ overflow: "hidden", pointerEvents: "none" }}>
                  <TextContent body={cell.text} w={cw} h={ch} theme={theme} />
                </foreignObject>
              )}
            </g>
          );
        }),
      )}
      {borderMode === "outside" && <rect width={w} height={h} fill="none" stroke={border} strokeWidth={1.4} />}
    </g>
  );
}

// ---------------- charts ----------------

function niceMax(v: number): number {
  if (v <= 0) return 1;
  const pow = Math.pow(10, Math.floor(Math.log10(v)));
  for (const m of [1, 2, 2.5, 5, 10]) {
    if (m * pow >= v) return m * pow;
  }
  return 10 * pow;
}

const AXIS_CLR = "#8C8C8C";
const GRID_CLR = "#E0E0E0";
const BASE_LABEL_STYLE: React.CSSProperties = { fontFamily: "Calibri, Arial, sans-serif", fontSize: 11, fill: AXIS_CLR };

/** Per-category cumulative bases for stacked variants; values scaled for percentStacked. */
function stackInfo(shape: ChartShape) {
  const grouping = shape.grouping ?? "clustered";
  const stacked = grouping === "stacked" || grouping === "percentStacked";
  const pct = grouping === "percentStacked";
  const nCats = shape.categories.length;
  const sums = Array.from({ length: nCats }, (_, ci) =>
    shape.series.reduce((a, s) => a + Math.max(0, s.values[ci] ?? 0), 0) || 1);
  const val = (si: number, ci: number) => {
    const v = Math.max(0, shape.series[si].values[ci] ?? 0);
    return pct ? (v / sums[ci]) * 100 : v;
  };
  const base = (si: number, ci: number) => {
    let b = 0;
    for (let k = 0; k < si; k++) b += val(k, ci);
    return b;
  };
  const maxV = pct ? 100 : stacked
    ? niceMax(Math.max(...sums))
    : niceMax(Math.max(1e-9, ...shape.series.flatMap(s => s.values.map(v => Math.abs(v)))));
  return { stacked, pct, val, base, maxV };
}

/** Straight or midpoint-smoothed path through points. */
function linePath(pts: (readonly [number, number])[], smooth?: boolean): string {
  if (!pts.length) return "";
  if (!smooth || pts.length < 3) return "M" + pts.map(p => `${p[0]} ${p[1]}`).join("L");
  let d = `M${pts[0][0]} ${pts[0][1]}`;
  for (let i = 1; i < pts.length - 1; i++) {
    const mx = (pts[i][0] + pts[i + 1][0]) / 2, my = (pts[i][1] + pts[i + 1][1]) / 2;
    d += `Q${pts[i][0]} ${pts[i][1]} ${mx} ${my}`;
  }
  d += `L${pts[pts.length - 1][0]} ${pts[pts.length - 1][1]}`;
  return d;
}

/** Margins + plot rectangle for a chart, in shape-local px. Shared by the renderer and the canvas (part hit/highlight). */
export function chartLayout(shape: ChartShape) {
  const w = px(shape.w), h = px(shape.h);
  const isPie = shape.chart === "pie" || shape.chart === "doughnut";
  const isRadar = shape.chart === "radar";
  const plain = isPie || isRadar;
  const hideX = !!shape.hideAxisX, hideY = !!shape.hideAxisY;
  const legendPos = shape.legend ? (shape.legendPos ?? "r") : null;
  const xTitleH = shape.axisTitleX && !plain ? 16 : 0;
  const yTitleW = shape.axisTitleY && !plain ? 16 : 0;
  const titleH = shape.title ? 26 : 8;
  const mTop = titleH + (legendPos === "t" ? 20 : 0);
  const mLeft = (legendPos === "l" ? 100 : 0) + yTitleW + (plain ? 8 : hideY ? 8 : shape.chart === "bar" ? 78 : 38);
  const mRight = (legendPos === "r" ? 96 : 6) + 2;
  const mBottom = (legendPos === "b" ? 20 : 0) + xTitleH + (plain ? 8 : hideX ? 8 : 22) + 4;
  const pw = Math.max(20, w - mLeft - mRight);
  const ph = Math.max(20, h - mTop - mBottom);
  return { w, h, plain, hideX, hideY, legendPos, xTitleH, yTitleW, titleH, ox: mLeft, oy: mTop, pw, ph };
}

/** Approximate clickable/highlightable boxes for each stylable chart part, in shape-local px. */
export function chartPartRegions(shape: ChartShape): { part: ChartPart; x: number; y: number; w: number; h: number }[] {
  const L = chartLayout(shape);
  const out: { part: ChartPart; x: number; y: number; w: number; h: number }[] = [];
  if (shape.title) out.push({ part: "title", x: L.w / 2 - 80, y: 2, w: 160, h: 22 });
  if (shape.axisTitleX && !L.plain) out.push({ part: "axisTitleX", x: L.ox + L.pw / 2 - 70, y: L.h - 20, w: 140, h: 16 });
  if (shape.axisTitleY && !L.plain) out.push({ part: "axisTitleY", x: (L.legendPos === "l" ? 104 : 4), y: L.oy + L.ph / 2 - 70, w: 16, h: 140 });
  if (L.legendPos) {
    if (L.legendPos === "r") out.push({ part: "legend", x: L.w - 96, y: L.oy, w: 92, h: Math.min(L.ph, 180) });
    else if (L.legendPos === "l") out.push({ part: "legend", x: 4, y: L.oy, w: 92, h: Math.min(L.ph, 180) });
    else if (L.legendPos === "t") out.push({ part: "legend", x: 6, y: L.titleH, w: L.w - 12, h: 18 });
    else out.push({ part: "legend", x: 6, y: L.h - 18, w: L.w - 12, h: 16 });
  }
  if (!L.plain) out.push({ part: "axisLabels", x: L.ox, y: L.oy + L.ph + 1, w: L.pw, h: 16 });
  return out;
}

export function ChartView({ shape, theme, defsPrefix = "c" }: { shape: ChartShape; theme: ColorTheme; defsPrefix?: string }) {
  const cats = shape.categories;
  const series = shape.series.length ? shape.series : [{ name: "", values: [] }];
  const isPie = shape.chart === "pie" || shape.chart === "doughnut";
  const stk = stackInfo(shape);

  const isRadar = shape.chart === "radar";
  const isScatter = shape.chart === "scatter";
  const { w, h, plain, hideX, hideY, legendPos, xTitleH, titleH, ox, oy, pw, ph } = chartLayout(shape);
  // chart-wide text size override (axis labels, legend, category names)
  const LABEL_STYLE: React.CSSProperties = {
    ...BASE_LABEL_STYLE,
    fontSize: shape.labelSizePt ? ptToPx(shape.labelSizePt) : ptToPx(12), // chart default: 12pt
  };
  // per-element style: merge a part's overrides over a fallback base style
  const partCss = (part: ChartPart, fallback: React.CSSProperties): React.CSSProperties => {
    const ps = shape.partStyles?.[part];
    if (!ps) return fallback;
    const fam = ps.font ? resolveFontName(ps.font, theme) : undefined;
    return {
      ...fallback,
      ...(fam ? { fontFamily: `'${fam}', Calibri, Arial, sans-serif` } : {}),
      ...(ps.sizePt ? { fontSize: ptToPx(ps.sizePt) } : {}),
      ...(ps.color ? { fill: resolveColor(ps.color, theme) } : {}),
      ...(ps.bold !== undefined ? { fontWeight: ps.bold ? "bold" : "normal" } : {}),
      ...(ps.italic ? { fontStyle: "italic" } : {}),
      ...(ps.underline ? { textDecoration: "underline" } : {}),
    };
  };
  const xs = isScatter
    ? cats.map((c, i) => { const n = parseFloat(c); return Number.isFinite(n) ? n : i + 1; })
    : [];
  const xMax = isScatter ? niceMax(Math.max(1e-9, ...xs.map(v => Math.abs(v)))) : 0;

  const maxV = stk.maxV;
  const ticks = [0, 0.25, 0.5, 0.75, 1].map(t => t * maxV);

  const sliceColor = (i: number) => seriesColor(i, theme, shape.pointColors?.[i] ?? undefined);
  const legendItems = isPie
    ? cats.map((c, i) => ({ label: c, color: sliceColor(i) }))
    : series.map((s, i) => ({ label: s.name, color: seriesColor(i, theme, s.color) }));

  const [bgPaint, bgDef] = paintFor(shape.chartFill, theme, `${defsPrefix}-cbg-${shape.id}`);
  const [plotPaint, plotDef] = paintFor(shape.plotFill, theme, `${defsPrefix}-pbg-${shape.id}`);
  const dl = !!shape.dataLabels;
  const err = shape.errorBarsPct;
  const dlPos = shape.dataLabelPos;                   // undefined = per-type default
  const DLBL = partCss("dataLabels", { ...LABEL_STYLE, fill: "#595959" });
  const AXLBL = partCss("axisLabels", LABEL_STYLE);   // axis tick + category labels
  const LEGEND = partCss("legend", LABEL_STYLE);      // legend entry text

  const dashArray = (d?: import("../model/types").LineDash) =>
    d === "dot" || d === "sysDot" ? "2 3"
    : d === "dashDot" ? "8 3 2 3"
    : d === "lgDash" ? "12 4"
    : d === "lgDashDot" ? "12 4 2 4"
    : d === "dash" ? "6 4"
    : undefined;
  const gridsOn = !shape.hideGridlines;
  const gridStroke = shape.gridColor ? resolveColor(shape.gridColor, theme) : GRID_CLR;
  // no chart-space outline unless the host/user sets one (PowerPoint has none by default)
  const chartStrokeOff = shape.chartBorder?.fill.kind === "none";
  const hasChartBorder = !!shape.chartBorder && !chartStrokeOff;
  const chartStroke = hasChartBorder ? resolveFill(shape.chartBorder!.fill, theme) : "none";
  const seriesLineW = (s: { lineWidthPt?: number }, fallback: number) =>
    s.lineWidthPt ? ptToPx(s.lineWidthPt) : fallback;
  const markerR = (fallback: number) => (shape.markerSizePt ? ptToPx(shape.markerSizePt) / 2 : fallback);

  /** error bar I-beam at a chart point. */
  const errBar = (key: React.Key, xPos: number, yPos: number, v: number, pxPerUnit: number, horiz = false) => {
    if (!err || v <= 0) return null;
    const e = ((v * err) / 100) * pxPerUnit;
    return horiz ? (
      <g key={key} stroke="#7F7F7F" strokeWidth={1}>
        <line x1={xPos - e} y1={yPos} x2={xPos + e} y2={yPos} />
        <line x1={xPos - e} y1={yPos - 3} x2={xPos - e} y2={yPos + 3} />
        <line x1={xPos + e} y1={yPos - 3} x2={xPos + e} y2={yPos + 3} />
      </g>
    ) : (
      <g key={key} stroke="#7F7F7F" strokeWidth={1}>
        <line x1={xPos} y1={yPos - e} x2={xPos} y2={yPos + e} />
        <line x1={xPos - 3} y1={yPos - e} x2={xPos + 3} y2={yPos - e} />
        <line x1={xPos - 3} y1={yPos + e} x2={xPos + 3} y2={yPos + e} />
      </g>
    );
  };

  return (
    <g transform={shapeTransform(shape)}>
      {bgDef && <defs>{bgDef}</defs>}
      {plotDef && <defs>{plotDef}</defs>}
      <rect
        width={w} height={h}
        fill={shape.chartFill ? (shape.chartFill.kind === "none" ? "none" : bgPaint) : "#FFFFFF"}
        stroke={chartStroke}
        strokeWidth={hasChartBorder ? ptToPx(shape.chartBorder!.widthPt) : 0}
        strokeDasharray={dashArray(shape.chartBorder?.dash)}
      />
      {!plain && (shape.plotFill || shape.plotBorder) && (
        <rect
          x={ox} y={oy} width={pw} height={ph}
          fill={shape.plotFill ? (shape.plotFill.kind === "none" ? "none" : plotPaint) : "none"}
          stroke={shape.plotBorder && shape.plotBorder.fill.kind !== "none" ? resolveFill(shape.plotBorder.fill, theme) : "none"}
          strokeWidth={shape.plotBorder ? ptToPx(shape.plotBorder.widthPt) : 0}
          strokeDasharray={dashArray(shape.plotBorder?.dash)}
        />
      )}
      {shape.title && (
        <text x={w / 2} y={18} textAnchor="middle" style={partCss("title", { ...LABEL_STYLE, fontSize: ptToPx(16), fill: "#404040" })}>{shape.title}</text>
      )}
      {shape.axisTitleX && !plain && (
        <text x={ox + pw / 2} y={h - (legendPos === "b" ? 24 : 5)} textAnchor="middle" style={partCss("axisTitleX", { ...LABEL_STYLE, fill: "#595959" })}>{shape.axisTitleX}</text>
      )}
      {shape.axisTitleY && !plain && (
        <text
          x={legendPos === "l" ? 112 : 12} y={oy + ph / 2}
          textAnchor="middle"
          transform={`rotate(-90 ${legendPos === "l" ? 112 : 12} ${oy + ph / 2})`}
          style={partCss("axisTitleY", { ...LABEL_STYLE, fill: "#595959" })}
        >{shape.axisTitleY}</text>
      )}

      {!isPie && !isRadar && (
        <g>
          {ticks.map((t, i) => {
            const frac = t / maxV;
            if (shape.chart === "bar") {
              const x = ox + frac * pw;
              return (
                <g key={i}>
                  {(gridsOn || i === 0) && <line x1={x} y1={oy} x2={x} y2={oy + ph} stroke={i === 0 ? GRID_CLR : gridStroke} strokeWidth={i === 0 ? 1.2 : 0.7} />}
                  {!hideX && <text x={x} y={oy + ph + 14} textAnchor="middle" style={AXLBL}>{fmtNum(t)}{stk.pct ? "%" : ""}</text>}
                </g>
              );
            }
            const y = oy + ph - frac * ph;
            return (
              <g key={i}>
                {(gridsOn || i === 0) && <line x1={ox} y1={y} x2={ox + pw} y2={y} stroke={i === 0 ? GRID_CLR : gridStroke} strokeWidth={i === 0 ? 1.2 : 0.7} />}
                {!hideY && <text x={ox - 5} y={y + 3.5} textAnchor="end" style={AXLBL}>{fmtNum(t)}{stk.pct ? "%" : ""}</text>}
              </g>
            );
          })}
          {isScatter && [0.25, 0.5, 0.75, 1].map(f => (
            <g key={`x${f}`}>
              {gridsOn && <line x1={ox + f * pw} y1={oy} x2={ox + f * pw} y2={oy + ph} stroke={gridStroke} strokeWidth={0.7} />}
              <text x={ox + f * pw} y={oy + ph + 14} textAnchor="middle" style={AXLBL}>{fmtNum(f * xMax)}</text>
            </g>
          ))}
        </g>
      )}

      {shape.chart === "column" && cats.map((cat, ci) => {
        const groupW = pw / cats.length;
        const barW = stk.stacked ? groupW * 0.55 : (groupW * 0.7) / series.length;
        return (
          <g key={ci}>
            {series.map((s, si) => {
              const v = stk.val(si, ci);
              const b = stk.stacked ? stk.base(si, ci) : 0;
              const bh = (v / maxV) * ph;
              const x = stk.stacked
                ? ox + ci * groupW + (groupW - barW) / 2
                : ox + ci * groupW + groupW * 0.15 + si * barW;
              const y = oy + ph - ((b + v) / maxV) * ph;
              const bw = Math.max(1, barW - (stk.stacked ? 0 : 1.5));
              return (
                <g key={si}>
                  <rect x={x} y={y} width={bw} height={bh} fill={seriesColor(si, theme, s.color)} stroke={stk.stacked ? "#FFFFFF" : undefined} strokeWidth={stk.stacked ? 0.8 : 0} />
                  {dl && v > 0 && (
                    <text
                      x={x + bw / 2}
                      y={stk.stacked || dlPos === "ctr" ? y + bh / 2 + 3
                        : dlPos === "inEnd" && bh >= 13 ? y + 11
                        : y - 3}
                      textAnchor="middle" style={DLBL}
                    >{fmtNum(s.values[ci] ?? 0)}</text>
                  )}
                  {errBar(`e${si}`, x + bw / 2, y, v, ph / maxV)}
                </g>
              );
            })}
            {!hideX && <text x={ox + ci * groupW + groupW / 2} y={oy + ph + 14} textAnchor="middle" style={AXLBL}>{cat}</text>}
          </g>
        );
      })}

      {shape.chart === "bar" && cats.map((cat, ci) => {
        const groupH = ph / cats.length;
        const barH = stk.stacked ? groupH * 0.55 : (groupH * 0.7) / series.length;
        return (
          <g key={ci}>
            {series.map((s, si) => {
              const v = stk.val(si, ci);
              const b = stk.stacked ? stk.base(si, ci) : 0;
              const bw = (v / maxV) * pw;
              const y = stk.stacked
                ? oy + ci * groupH + (groupH - barH) / 2
                : oy + ci * groupH + groupH * 0.15 + si * barH;
              const bh = Math.max(1, barH - (stk.stacked ? 0 : 1.5));
              const xEnd = ox + ((b + v) / maxV) * pw;
              return (
                <g key={si}>
                  <rect x={ox + (b / maxV) * pw} y={y} width={bw} height={bh} fill={seriesColor(si, theme, s.color)} stroke={stk.stacked ? "#FFFFFF" : undefined} strokeWidth={stk.stacked ? 0.8 : 0} />
                  {dl && v > 0 && (() => {
                    const center = stk.stacked || dlPos === "ctr";
                    const inEnd = !center && dlPos === "inEnd" && bw >= 22;
                    const lx = center ? xEnd - bw / 2 : inEnd ? xEnd - 4 : xEnd + 4;
                    return <text x={lx} y={y + bh / 2 + 3.5} textAnchor={center ? "middle" : inEnd ? "end" : "start"} style={DLBL}>{fmtNum(s.values[ci] ?? 0)}</text>;
                  })()}
                  {errBar(`e${si}`, xEnd, y + bh / 2, v, pw / maxV, true)}
                </g>
              );
            })}
            {!hideY && <text x={ox - 6} y={oy + ci * groupH + groupH / 2 + 3.5} textAnchor="end" style={AXLBL}>{truncate(cat, 14)}</text>}
          </g>
        );
      })}

      {(shape.chart === "line" || shape.chart === "area") && (
        <g>
          {series.map((s, si) => {
            const pts = cats.map((_, ci) => {
              const v = stk.stacked ? stk.base(si, ci) + stk.val(si, ci) : Math.max(0, s.values[ci] ?? 0);
              const x = ox + ((ci + 0.5) / cats.length) * pw;
              const y = oy + ph - (v / maxV) * ph;
              return [x, y] as const;
            });
            const basePts = cats.map((_, ci) => {
              const v = stk.stacked ? stk.base(si, ci) : 0;
              return [ox + ((ci + 0.5) / cats.length) * pw, oy + ph - (v / maxV) * ph] as const;
            });
            const color = seriesColor(si, theme, s.color);
            return (
              <g key={si}>
                {shape.chart === "area" && pts.length > 0 && (
                  <path
                    d={`${linePath(pts)}L${[...basePts].reverse().map(p => `${p[0]} ${p[1]}`).join("L")}Z`}
                    fill={color} opacity={0.4}
                  />
                )}
                <path d={linePath(pts, shape.smooth)} fill="none" stroke={color} strokeWidth={seriesLineW(s, 2.2)} strokeDasharray={dashArray(s.dash)} strokeLinejoin="round" />
                {shape.chart === "line" && (shape.marker ?? true) && pts.map((p, i) => <circle key={i} cx={p[0]} cy={p[1]} r={markerR(2.6)} fill={color} />)}
                {dl && pts.map((p, i) => (
                  <text key={`l${i}`} x={p[0]} y={dlPos === "ctr" ? p[1] + 3 : dlPos === "inEnd" ? p[1] + 14 : p[1] - 6} textAnchor="middle" style={DLBL}>{fmtNum(s.values[i] ?? 0)}</text>
                ))}
                {err && pts.map((p, i) => errBar(`e${i}`, p[0], p[1], stk.stacked ? stk.base(si, i) + stk.val(si, i) : Math.max(0, s.values[i] ?? 0), ph / maxV))}
              </g>
            );
          })}
          {!hideX && cats.map((cat, ci) => (
            <text key={ci} x={ox + ((ci + 0.5) / cats.length) * pw} y={oy + ph + 14} textAnchor="middle" style={AXLBL}>{cat}</text>
          ))}
        </g>
      )}

      {isScatter && (
        <g>
          {series.map((s, si) => {
            const color = seriesColor(si, theme, s.color);
            const pts = xs.map((xv, ci) => [
              ox + (Math.max(0, xv) / xMax) * pw,
              oy + ph - (Math.max(0, s.values[ci] ?? 0) / maxV) * ph,
            ] as const).sort((a, b) => a[0] - b[0]);
            return (
              <g key={si}>
                {/* smooth === undefined -> markers only; false -> straight lines; true -> smooth */}
                {shape.smooth !== undefined && (
                  <path d={linePath(pts, shape.smooth)} fill="none" stroke={color} strokeWidth={seriesLineW(s, 2)} strokeDasharray={dashArray(s.dash)} strokeLinejoin="round" />
                )}
                {(shape.marker ?? true) && pts.map((p, i) => <circle key={i} cx={p[0]} cy={p[1]} r={markerR(3)} fill={color} />)}
                {dl && pts.map((p, i) => (
                  <text key={`l${i}`} x={p[0]} y={dlPos === "ctr" ? p[1] + 3 : dlPos === "inEnd" ? p[1] + 14 : p[1] - 7} textAnchor="middle" style={DLBL}>{fmtNum(s.values[i] ?? 0)}</text>
                ))}
                {err && pts.map((p, i) => errBar(`e${i}`, p[0], p[1], Math.max(0, s.values[i] ?? 0), ph / maxV))}
              </g>
            );
          })}
        </g>
      )}

      {isRadar && (() => {
        const cx = ox + pw / 2, cy = oy + ph / 2;
        const R = Math.min(pw, ph) / 2 - 14;
        const n = Math.max(cats.length, 3);
        const angle = (i: number) => -Math.PI / 2 + (i / n) * Math.PI * 2;
        const pt = (i: number, frac: number) => [cx + R * frac * Math.cos(angle(i)), cy + R * frac * Math.sin(angle(i))] as const;
        return (
          <g>
            {gridsOn && [0.25, 0.5, 0.75, 1].map(f => (
              <polygon key={f} points={Array.from({ length: n }, (_, i) => pt(i, f).join(",")).join(" ")} fill="none" stroke={gridStroke} strokeWidth={0.8} />
            ))}
            {gridsOn && Array.from({ length: n }, (_, i) => (
              <line key={i} x1={cx} y1={cy} x2={pt(i, 1)[0]} y2={pt(i, 1)[1]} stroke={gridStroke} strokeWidth={0.8} />
            ))}
            {cats.map((c, i) => {
              const [lx, ly] = pt(i, 1.13);
              return <text key={i} x={lx} y={ly + 3} textAnchor="middle" style={AXLBL}>{truncate(c, 8)}</text>;
            })}
            {series.map((s, si) => {
              const color = seriesColor(si, theme, s.color);
              const poly = Array.from({ length: n }, (_, i) => pt(i, Math.max(0, s.values[i] ?? 0) / maxV));
              const points = poly.map(p => p.join(",")).join(" ");
              const filled = shape.radarStyle === "filled";
              return (
                <g key={si}>
                  <polygon points={points} fill={filled ? color : "none"} opacity={filled ? 0.45 : 1} stroke={color} strokeWidth={seriesLineW(s, 2)} strokeDasharray={dashArray(s.dash)} strokeLinejoin="round" />
                  {shape.radarStyle === "marker" && poly.map((p, i) => <circle key={i} cx={p[0]} cy={p[1]} r={markerR(2.6)} fill={color} />)}
                </g>
              );
            })}
          </g>
        );
      })()}

      {isPie && (() => {
        const vals = (series[0]?.values ?? []).map(v => Math.max(0, v));
        const total = vals.reduce((a, b) => a + b, 0) || 1;
        const cx = ox + pw / 2, cy = oy + ph / 2;
        const R = Math.min(pw, ph) / 2 - 4;
        const r0 = shape.chart === "doughnut" ? R * 0.55 : 0;
        const outside = dlPos === "outEnd";
        const pieDlbl = partCss("dataLabels", { ...LABEL_STYLE, fill: outside ? "#595959" : "#FFFFFF" });
        let ang = -Math.PI / 2;
        return vals.map((v, i) => {
          const sweep = (v / total) * Math.PI * 2;
          const a0 = ang, a1 = ang + sweep;
          ang = a1;
          const large = sweep > Math.PI ? 1 : 0;
          const p = (a: number, rad: number) => `${cx + rad * Math.cos(a)} ${cy + rad * Math.sin(a)}`;
          const d = r0 > 0
            ? `M${p(a0, R)}A${R} ${R} 0 ${large} 1 ${p(a1, R)}L${p(a1, r0)}A${r0} ${r0} 0 ${large} 0 ${p(a0, r0)}Z`
            : `M${cx} ${cy}L${p(a0, R)}A${R} ${R} 0 ${large} 1 ${p(a1, R)}Z`;
          const mid = (a0 + a1) / 2;
          const lr = outside ? R + 12
            : dlPos === "inEnd" ? (r0 > 0 ? R * 0.78 + r0 * 0.22 : R * 0.82)
            : (r0 > 0 ? (R + r0) / 2 : R * 0.62);   // ctr / default
          return (
            <g key={i}>
              <path d={d} fill={sliceColor(i)} stroke="#FFFFFF" strokeWidth={1.2} />
              {dl && v > 0 && sweep > 0.12 && (
                <text
                  x={cx + lr * Math.cos(mid)} y={cy + lr * Math.sin(mid) + 3}
                  textAnchor={outside ? (Math.cos(mid) >= 0 ? "start" : "end") : "middle"}
                  style={pieDlbl}
                >{Math.round((v / total) * 100)}%</text>
              )}
            </g>
          );
        });
      })()}

      {legendPos && (() => {
        const items = legendItems.slice(0, 12);
        const fontPx = typeof LEGEND.fontSize === "number" ? LEGEND.fontSize : 12;
        // chars that fit a text budget at this font (~0.58em average glyph width)
        const fit = (budgetPx: number) => Math.max(1, Math.floor(budgetPx / (fontPx * 0.58)));
        const swatch = 9, gap = 4;
        if (legendPos === "t" || legendPos === "b") {
          // shrink each slot to fit the chart width — never overflow the bounds
          const itemW = Math.min(112, (w - 12) / items.length);
          const x0 = Math.max(6, (w - itemW * items.length) / 2);
          const y = legendPos === "t" ? titleH + 10 : h - 8;
          return (
            <g transform={`translate(${x0} ${y})`}>
              {items.map((it, i) => (
                <g key={i} transform={`translate(${i * itemW} 0)`}>
                  <rect width={swatch} height={swatch} y={-swatch} fill={it.color} />
                  <text x={swatch + gap} y={0} style={LEGEND}>{truncate(it.label, fit(itemW - swatch - gap - 2))}</text>
                </g>
              ))}
            </g>
          );
        }
        const boxW = 96;
        const x = legendPos === "l" ? 8 : w - boxW + 8;
        const rowH = Math.max(16, fontPx + 5);
        return (
          <g transform={`translate(${x} ${oy + 6})`}>
            {items.map((it, i) => (
              <g key={i} transform={`translate(0 ${i * rowH})`}>
                <rect width={swatch} height={swatch} y={-swatch} fill={it.color} />
                <text x={swatch + gap} y={0} style={LEGEND}>{truncate(it.label, fit(boxW - swatch - gap - 8))}</text>
              </g>
            ))}
          </g>
        );
      })()}
    </g>
  );
}

function fmtNum(v: number): string {
  if (Math.abs(v) >= 1000) return `${Math.round(v / 100) / 10}k`;
  return `${Math.round(v * 100) / 100}`;
}

function truncate(s: string, n: number): string {
  return s.length > n ? s.slice(0, n - 1) + "…" : s;
}
