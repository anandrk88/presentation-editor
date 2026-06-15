import JSZip from "jszip";
import type {
  ChartShape, ColorRef, Fill, LineProps, MediaItem, Paragraph, PicShape,
  Presentation, Run, Shape, SlideModel, SpShape, TableShape, TextBody,
} from "../model/types";
import { EMU_PER_PT, resolveColor, tableBorderColor, tableCellStyle } from "../model/defaults";
import { cssColorToHex, tintSvgText } from "../util/svgTint";

/**
 * Serializes the document model into a complete OOXML PresentationML package
 * (ECMA-376): content types, package/part relationships, presentation, theme,
 * slide master, slide layout, slides and media. Output opens in PowerPoint,
 * LibreOffice and other OOXML suites.
 */

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;

export function esc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function colorXml(c: ColorRef): string {
  let mods = "";
  if (c.alpha !== undefined && c.alpha < 100) mods += `<a:alpha val="${Math.round(c.alpha * 1000)}"/>`;
  if (c.lumMod !== undefined) mods += `<a:lumMod val="${Math.round(c.lumMod * 1000)}"/>`;
  if (c.lumOff !== undefined) mods += `<a:lumOff val="${Math.round(c.lumOff * 1000)}"/>`;
  const tag = c.kind === "srgb" ? "a:srgbClr" : "a:schemeClr";
  const val = c.kind === "srgb" ? c.hex : c.slot;
  return mods ? `<${tag} val="${val}">${mods}</${tag}>` : `<${tag} val="${val}"/>`;
}

function fillXml(f: Fill, imageRels?: Map<string, string>): string {
  if (f.kind === "none") return `<a:noFill/>`;
  if (f.kind === "gradient") {
    const gs = [...f.stops].sort((a, b) => a.pos - b.pos)
      .map(s => `<a:gs pos="${Math.round(s.pos * 1000)}">${colorXml(s.color)}</a:gs>`).join("");
    return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst><a:lin ang="${Math.round(f.angle * 60000)}" scaled="1"/></a:gradFill>`;
  }
  if (f.kind === "image") {
    const rid = imageRels?.get(f.mediaId);
    if (!rid) return `<a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill>`; // media missing
    const mode = f.tile
      ? `<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>`
      : `<a:stretch><a:fillRect/></a:stretch>`;
    return `<a:blipFill rotWithShape="1"><a:blip r:embed="${rid}"/>${mode}</a:blipFill>`;
  }
  if (f.kind === "pattern") {
    return `<a:pattFill prst="${esc(f.prst)}"><a:fgClr>${colorXml(f.fg)}</a:fgClr><a:bgClr>${colorXml(f.bg)}</a:bgClr></a:pattFill>`;
  }
  return `<a:solidFill>${colorXml(f.color)}</a:solidFill>`;
}

function arrowEndXml(tag: "headEnd" | "tailEnd", e: import("../model/types").ArrowEnd | undefined): string {
  if (!e || e.type === "none") return "";
  return `<a:${tag} type="${e.type}" w="${e.w ?? "med"}" len="${e.len ?? "med"}"/>`;
}

function lineXml(l: LineProps): string {
  const w = Math.max(1, Math.round(l.widthPt * EMU_PER_PT));
  if (l.fill.kind === "none") return `<a:ln><a:noFill/></a:ln>`;
  const dash = l.dash && l.dash !== "solid" ? `<a:prstDash val="${l.dash}"/>` : "";
  // headEnd/tailEnd come after dash + line-join in CT_LineProperties
  const arrows = arrowEndXml("headEnd", l.headEnd) + arrowEndXml("tailEnd", l.tailEnd);
  return `<a:ln w="${w}">${fillXml(l.fill)}${dash}<a:round/>${arrows}</a:ln>`;
}

function xfrmXml(s: Shape): string {
  const rot = s.rot ? ` rot="${Math.round(s.rot * 60000)}"` : "";
  const fh = s.flipH ? ` flipH="1"` : "";
  const fv = s.flipV ? ` flipV="1"` : "";
  return `<a:xfrm${rot}${fh}${fv}><a:off x="${Math.round(s.x)}" y="${Math.round(s.y)}"/><a:ext cx="${Math.max(0, Math.round(s.w))}" cy="${Math.max(0, Math.round(s.h))}"/></a:xfrm>`;
}

function runPrXml(r: Run, tag: "a:rPr" | "a:endParaRPr"): string {
  const attrs = [
    `lang="en-US"`,
    `sz="${Math.round(r.sizePt * 100)}"`,
    r.bold ? `b="1"` : "",
    r.italic ? `i="1"` : "",
    r.underline ? `u="sng"` : "",
    r.strike ? `strike="sngStrike"` : "",
    r.baseline ? `baseline="${Math.round(r.baseline * 1000)}"` : "",
    r.caps ? `cap="${r.caps}"` : "",
    `dirty="0"`,
  ].filter(Boolean).join(" ");
  const highlight = r.highlight ? `<a:highlight>${colorXml(r.highlight)}</a:highlight>` : "";
  return `<${tag} ${attrs}><a:solidFill>${colorXml(r.color)}</a:solidFill>${highlight}<a:latin typeface="${esc(r.font)}"/></${tag}>`;
}

function paraXml(p: Paragraph, fallback: Run): string {
  const indentEmu = 342900; // 0.375" hanging indent for bullets
  const marL = p.level * 457200 + (p.bullet !== "none" ? indentEmu : 0);
  const attrs = [
    marL ? `marL="${marL}"` : "",
    p.bullet !== "none" ? `indent="-${indentEmu}"` : "",
    p.level ? `lvl="${p.level}"` : "",
    `algn="${p.align}"`,
  ].filter(Boolean).join(" ");
  let children = "";
  if (p.lineSpacingPct) children += `<a:lnSpc><a:spcPct val="${Math.round(p.lineSpacingPct * 1000)}"/></a:lnSpc>`;
  if (p.bullet === "none") children += `<a:buNone/>`;
  else if (p.bullet === "char") children += `<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/>`;
  else children += `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>`;
  const pPr = `<a:pPr ${attrs}>${children}</a:pPr>`;
  const runs = p.runs.map(r => `<a:r>${runPrXml(r, "a:rPr")}<a:t>${esc(r.text)}</a:t></a:r>`).join("");
  const end = runPrXml(p.runs.length ? p.runs[p.runs.length - 1] : fallback, "a:endParaRPr");
  return `<a:p>${pPr}${runs}${end}</a:p>`;
}

function txBodyXml(t: TextBody | undefined, fallback: Run, tag: "p:txBody" | "a:txBody" = "p:txBody"): string {
  if (!t) {
    return `<${tag}><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>${runPrXml(fallback, "a:endParaRPr")}</a:p></${tag}>`;
  }
  const [l, tp, r, b] = t.insets.map(Math.round);
  const colAttrs = t.columns && t.columns > 1
    ? ` numCol="${t.columns}" spcCol="${Math.round(t.colSpacing ?? 360000)}"`
    : "";
  const bodyPr = `<a:bodyPr rtlCol="0" wrap="${t.wrap ? "square" : "none"}" lIns="${l}" tIns="${tp}" rIns="${r}" bIns="${b}"${colAttrs} anchor="${t.anchor}"/>`;
  const paras = t.paragraphs.map(p => paraXml(p, fallback)).join("");
  return `<${tag}>${bodyPr}<a:lstStyle/>${paras || `<a:p>${runPrXml(fallback, "a:endParaRPr")}</a:p>`}</${tag}>`;
}

const DEFAULT_RUN: Run = { text: "", sizePt: 18, font: "+mn-lt", color: { kind: "scheme", slot: "dk1" } };

function avLstXml(adj?: Record<string, number>): string {
  return adj
    ? Object.entries(adj).map(([n, v]) => `<a:gd name="${esc(n)}" fmla="val ${Math.round(v)}"/>`).join("")
    : "";
}

function spXml(s: SpShape, numId: number, imageRels?: Map<string, string>): string {
  const txBox = s.isTextBox ? ` txBox="1"` : "";
  const geom = s.rawGeomXml
    ? s.rawGeomXml // imported freeform, preserved verbatim
    : `<a:prstGeom prst="${esc(s.geom)}"><a:avLst>${avLstXml(s.adj)}</a:avLst></a:prstGeom>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${numId}" name="${esc(s.name)}"/><p:cNvSpPr${txBox}/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrmXml(s)}${geom}${fillXml(s.fill, imageRels)}${lineXml(s.line)}</p:spPr>` +
    txBodyXml(s.text, DEFAULT_RUN) +
    `</p:sp>`;
}

// ---------- tables ----------
function tableXml(s: TableShape, numId: number): string {
  const wSum = s.colW.reduce((a, b) => a + b, 0) || 1;
  const hSum = s.rowH.reduce((a, b) => a + b, 0) || 1;
  const colW = s.colW.map(c => Math.max(1, Math.round((c / wSum) * s.w)));
  const rowH = s.rowH.map(r => Math.max(1, Math.round((r / hSum) * s.h)));
  const grid = colW.map(wv => `<a:gridCol w="${wv}"/>`).join("");

  const borderMode = s.borderMode ?? "all";
  const borderClr = colorXml(tableBorderColor(s));
  const lnOn = (tag: string) =>
    `<a:${tag} w="12700" cap="flat"><a:solidFill>${borderClr}</a:solidFill><a:prstDash val="solid"/></a:${tag}>`;
  const lnOff = (tag: string) => `<a:${tag} w="12700"><a:noFill/></a:${tag}>`;
  const nRows = s.cells.length, nCols = s.cells[0]?.length ?? 0;
  // per-edge borders: "all" = every edge, "outside" = table boundary only, "none" = none
  const edges = (r: number, c: number, span: number, vspan: number) => {
    const on = (outer: boolean) => (borderMode === "all" ? true : borderMode === "outside" ? outer : false);
    return (on(c === 0) ? lnOn("lnL") : lnOff("lnL")) +
      (on(c + span >= nCols) ? lnOn("lnR") : lnOff("lnR")) +
      (on(r === 0) ? lnOn("lnT") : lnOff("lnT")) +
      (on(r + vspan >= nRows) ? lnOn("lnB") : lnOff("lnB"));
  };

  const rows = s.cells.map((row, r) => {
    const tcs = row.map((cell, c) => {
      if (cell.merged === "h") return `<a:tc hMerge="1">${txBodyXml(cell.text, DEFAULT_RUN, "a:txBody")}<a:tcPr/></a:tc>`;
      if (cell.merged === "v") return `<a:tc vMerge="1">${txBodyXml(cell.text, DEFAULT_RUN, "a:txBody")}<a:tcPr/></a:tc>`;
      const span = Math.max(1, cell.gridSpan ?? 1);
      const vspan = Math.max(1, cell.rowSpan ?? 1);
      const spanAttr = span > 1 ? ` gridSpan="${span}"` : "";
      const vspanAttr = vspan > 1 ? ` rowSpan="${vspan}"` : "";
      const [mL, mT, mR, mB] = cell.text.insets;
      const fill = cell.fill ?? tableCellStyle(s, r, c).fill;
      const tcPr = `<a:tcPr marL="${Math.round(mL)}" marR="${Math.round(mR)}" marT="${Math.round(mT)}" marB="${Math.round(mB)}" anchor="${cell.text.anchor}">` +
        edges(r, c, span, vspan) +
        fillXml(fill) + `</a:tcPr>`;
      return `<a:tc${spanAttr}${vspanAttr}>${txBodyXml(cell.text, DEFAULT_RUN, "a:txBody")}${tcPr}</a:tc>`;
    }).join("");
    return `<a:tr h="${rowH[r] ?? rowH[rowH.length - 1]}">${tcs}</a:tr>`;
  }).join("");

  const tblFlags = `firstRow="${s.firstRow ? 1 : 0}" lastRow="${s.totalRow ? 1 : 0}" firstCol="${s.firstCol ? 1 : 0}" lastCol="${s.lastCol ? 1 : 0}" bandRow="${s.bandRow ? 1 : 0}" bandCol="${s.bandCol ? 1 : 0}"`;
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${numId}" name="${esc(s.name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${Math.round(s.x)}" y="${Math.round(s.y)}"/><a:ext cx="${Math.max(1, Math.round(s.w))}" cy="${Math.max(1, Math.round(s.h))}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
    `<a:tbl><a:tblPr ${tblFlags}><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>` +
    `<a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl>` +
    `</a:graphicData></a:graphic></p:graphicFrame>`;
}

// ---------- charts ----------
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const ACCENT_SLOTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

function chartColor(i: number, explicit?: ColorRef): ColorRef {
  return explicit ?? { kind: "scheme", slot: ACCENT_SLOTS[i % ACCENT_SLOTS.length] };
}

/** <a:defRPr>/<a:rPr> for a chart text element from a per-part style + fallback size. */
function chartRunPr(tag: "a:defRPr" | "a:rPr", style: import("../model/types").ChartTextStyle | undefined, fallbackSz?: number): string {
  const sz = style?.sizePt ?? fallbackSz;
  const attrs = [
    tag === "a:rPr" ? `lang="en-US"` : "",
    sz !== undefined ? `sz="${Math.round(sz * 100)}"` : "",
    style?.bold !== undefined ? `b="${style.bold ? 1 : 0}"` : "",
    style?.italic ? `i="1"` : "",
    style?.underline ? `u="sng"` : "",
  ].filter(Boolean).join(" ");
  const children =
    (style?.color ? `<a:solidFill>${colorXml(style.color)}</a:solidFill>` : "") +
    (style?.font ? `<a:latin typeface="${esc(style.font)}"/>` : "");
  return `<${tag}${attrs ? " " + attrs : ""}>${children}</${tag}>`;
}

/** <c:txPr> block (used by legend + axes) carrying a per-part style. */
function chartTxPr(style: import("../model/types").ChartTextStyle | undefined, fallbackSz?: number): string {
  if (!style && fallbackSz === undefined) return "";
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${chartRunPr("a:defRPr", style, fallbackSz)}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

function strLit(values: string[]): string {
  return `<c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("");
}

function numLit(values: number[]): string {
  return `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${Number.isFinite(v) ? v : 0}</c:v></c:pt>`).join("");
}

export function chartSpaceXml(ch: ChartShape): string {
  const isPie = ch.chart === "pie" || ch.chart === "doughnut";
  const isScatter = ch.chart === "scatter";
  const axId1 = 111111111, axId2 = 222222222;
  const grouping = ch.grouping ?? "clustered";
  const lineGrouping = ch.grouping && ch.grouping !== "clustered" ? ch.grouping : "standard";

  const markerSize = Math.max(2, Math.min(72, Math.round(ch.markerSizePt ?? 5)));
  const markerXml = (on: boolean) =>
    `<c:marker><c:symbol val="${on ? "circle" : "none"}"/>${on ? `<c:size val="${markerSize}"/>` : ""}</c:marker>`;

  const serXml = ch.series.map((s, i) => {
    // CT_SerTx only allows strRef | v — a strLit here makes PowerPoint repair-strip the chart
    const head = `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${esc(s.name)}</c:v></c:tx>`;
    const fill = `<a:solidFill>${colorXml(chartColor(i, s.color))}</a:solidFill>`;
    const serW = Math.round((s.lineWidthPt ?? 2.25) * 12700);
    const serDash = s.dash && s.dash !== "solid" ? `<a:prstDash val="${s.dash}"/>` : "";
    const lnOn = `<a:ln w="${serW}"><a:solidFill>${colorXml(chartColor(i, s.color))}</a:solidFill>${serDash}</a:ln>`;
    const lnOff = `<a:ln w="28575"><a:noFill/></a:ln>`;

    if (isScatter) {
      const xsVals = ch.categories.map((c, ci) => { const n = parseFloat(c); return Number.isFinite(n) ? n : ci + 1; });
      const lines = ch.smooth !== undefined;
      return head +
        `<c:spPr>${lines ? lnOn : lnOff}</c:spPr>` +
        markerXml(ch.marker ?? true) +
        `<c:xVal><c:numLit>${numLit(xsVals)}</c:numLit></c:xVal>` +
        `<c:yVal><c:numLit>${numLit(ch.categories.map((_, ci) => s.values[ci] ?? 0))}</c:numLit></c:yVal>` +
        `<c:smooth val="${ch.smooth ? 1 : 0}"/></c:ser>`;
    }

    const dPts = isPie
      ? ch.categories.map((_, ci) =>
          `<c:dPt><c:idx val="${ci}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${colorXml(chartColor(ci, ch.pointColors?.[ci] ?? undefined))}</a:solidFill></c:spPr></c:dPt>`).join("")
      : "";
    const marker = ch.chart === "line" ? markerXml(ch.marker ?? true)
      : ch.chart === "radar" ? markerXml(ch.radarStyle === "marker")
      : "";
    const smooth = ch.chart === "line" ? `<c:smooth val="${ch.smooth ? 1 : 0}"/>` : "";
    const spPr = `<c:spPr>${fill}${ch.chart === "line" || ch.chart === "radar" ? lnOn : ""}</c:spPr>`;
    const errBars = ch.errorBarsPct
      ? `<c:errBars><c:errBarType val="both"/><c:errValType val="percentage"/><c:noEndCap val="0"/><c:val val="${ch.errorBarsPct}"/></c:errBars>`
      : "";
    return head + spPr + marker + dPts + errBars +
      `<c:cat><c:strLit>${strLit(ch.categories)}</c:strLit></c:cat>` +
      `<c:val><c:numLit>${numLit(ch.categories.map((_, ci) => s.values[ci] ?? 0))}</c:numLit></c:val>` +
      smooth +
      `</c:ser>`;
  }).join("");

  const dLbls = ch.dataLabels
    ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${isPie ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${isPie ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`
    : "";

  const axTitle = (text: string | undefined, horz: boolean, style?: import("../model/types").ChartTextStyle) => text
    ? `<c:title><c:tx><c:rich><a:bodyPr${horz ? "" : ` rot="-5400000" vert="horz"`}/><a:lstStyle/><a:p><a:pPr>${chartRunPr("a:defRPr", style, 9)}</a:pPr><a:r>${chartRunPr("a:rPr", style, 9)}<a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
    : "";
  // horizontal (bottom) axis = cat for column/line/area, val for bar; swap for bar charts
  const isBar = ch.chart === "bar";
  const ps = ch.partStyles ?? {};
  const catTitle = axTitle(isBar ? ch.axisTitleY : ch.axisTitleX, !isBar, isBar ? ps.axisTitleY : ps.axisTitleX);
  const valTitle = axTitle(isBar ? ch.axisTitleX : ch.axisTitleY, isBar, isBar ? ps.axisTitleX : ps.axisTitleY);
  // per-axis label text style (written to both axes so cat + val labels match)
  const axLblTxPr = ps.axisLabels || ch.labelSizePt ? chartTxPr(ps.axisLabels, ch.labelSizePt) : "";
  const catDelete = isBar ? ch.hideAxisY : ch.hideAxisX;
  const valDelete = isBar ? ch.hideAxisX : ch.hideAxisY;

  // value-axis gridlines: omitted when hidden, styled <c:spPr> when recolored
  const grid = ch.hideGridlines
    ? ""
    : ch.gridColor
      ? `<c:majorGridlines><c:spPr><a:ln w="9525" cap="flat"><a:solidFill>${colorXml(ch.gridColor)}</a:solidFill><a:round/></a:ln></c:spPr></c:majorGridlines>`
      : `<c:majorGridlines/>`;

  let plot: string;
  const axes =
    `<c:catAx><c:axId val="${axId1}"/><c:scaling><c:orientation val="${ch.chart === "bar" ? "maxMin" : "minMax"}"/></c:scaling><c:delete val="${catDelete ? 1 : 0}"/><c:axPos val="${ch.chart === "bar" ? "l" : "b"}"/>${catTitle}${axLblTxPr}<c:crossAx val="${axId2}"/></c:catAx>` +
    `<c:valAx><c:axId val="${axId2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${valDelete ? 1 : 0}"/><c:axPos val="${ch.chart === "bar" ? "b" : "l"}"/>${grid}${valTitle}${axLblTxPr}<c:crossAx val="${axId1}"/></c:valAx>`;
  const scatterAxes =
    `<c:valAx><c:axId val="${axId1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${ch.hideAxisX ? 1 : 0}"/><c:axPos val="b"/>${axTitle(ch.axisTitleX, true, ps.axisTitleX)}${axLblTxPr}<c:crossAx val="${axId2}"/></c:valAx>` +
    `<c:valAx><c:axId val="${axId2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${ch.hideAxisY ? 1 : 0}"/><c:axPos val="l"/>${grid}${axTitle(ch.axisTitleY, false, ps.axisTitleY)}${axLblTxPr}<c:crossAx val="${axId1}"/></c:valAx>`;
  const axIds = `<c:axId val="${axId1}"/><c:axId val="${axId2}"/>`;
  const overlap = grouping === "stacked" || grouping === "percentStacked" ? `<c:overlap val="100"/>` : "";

  switch (ch.chart) {
    case "column":
    case "bar":
      plot = `<c:barChart><c:barDir val="${ch.chart === "bar" ? "bar" : "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${serXml}${dLbls}<c:gapWidth val="150"/>${overlap}${axIds}</c:barChart>${axes}`;
      break;
    case "line":
      plot = `<c:lineChart><c:grouping val="${lineGrouping}"/><c:varyColors val="0"/>${serXml}${dLbls}<c:marker val="1"/>${axIds}</c:lineChart>${axes}`;
      break;
    case "area":
      plot = `<c:areaChart><c:grouping val="${lineGrouping}"/><c:varyColors val="0"/>${serXml}${dLbls}${axIds}</c:areaChart>${axes}`;
      break;
    case "scatter":
      plot = `<c:scatterChart><c:scatterStyle val="${ch.smooth === undefined ? "marker" : ch.smooth ? "smoothMarker" : "lineMarker"}"/><c:varyColors val="0"/>${serXml}${dLbls}${axIds}</c:scatterChart>${scatterAxes}`;
      break;
    case "radar":
      plot = `<c:radarChart><c:radarStyle val="${ch.radarStyle ?? "standard"}"/><c:varyColors val="0"/>${serXml}${dLbls}${axIds}</c:radarChart>${axes}`;
      break;
    case "pie":
      plot = `<c:pieChart><c:varyColors val="1"/>${serXml}${dLbls}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case "doughnut":
      plot = `<c:doughnutChart><c:varyColors val="1"/>${serXml}${dLbls}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
      break;
  }

  const title = ch.title
    ? `<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr>${chartRunPr("a:defRPr", { bold: false, ...ps.title }, 14)}</a:pPr><a:r>${chartRunPr("a:rPr", { bold: false, ...ps.title }, 14)}<a:t>${esc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : `<c:autoTitleDeleted val="1"/>`;
  const legend = ch.legend ? `<c:legend><c:legendPos val="${ch.legendPos ?? "r"}"/><c:overlay val="0"/>${chartTxPr(ps.legend, ch.labelSizePt)}</c:legend>` : "";

  // plot area styling sits inside c:plotArea (after the chart group + axes)
  const plotSpPr = ch.plotFill || ch.plotBorder
    ? `<c:spPr>${ch.plotFill ? fillXml(ch.plotFill) : ""}${ch.plotBorder ? lineXml(ch.plotBorder) : ""}</c:spPr>`
    : "";
  const spPr = ch.chartFill || ch.chartBorder
    ? `<c:spPr>${ch.chartFill ? fillXml(ch.chartFill) : ""}${ch.chartBorder ? lineXml(ch.chartBorder) : ""}</c:spPr>`
    : "";
  const txPr = ch.labelSizePt
    ? `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round(ch.labelSizePt * 100)}"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`
    : "";

  return XML_DECL +
    `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    `<c:chart>${title}<c:plotArea><c:layout/>${plot}${plotSpPr}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    spPr + txPr +
    `</c:chartSpace>`;
}

function chartFrameXml(s: ChartShape, numId: number, rId: string): string {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${numId}" name="${esc(s.name)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${Math.round(s.x)}" y="${Math.round(s.y)}"/><a:ext cx="${Math.max(1, Math.round(s.w))}" cy="${Math.max(1, Math.round(s.h))}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${rId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

function picXml(s: PicShape, numId: number, refs: { rid?: string; svgRid?: string }): string {
  // vector images: asvg extension, with a bitmap fallback in the main blip when we
  // have one — or PowerPoint's pure-graphic form (blip without r:embed) when we don't
  const svgExt = refs.svgRid
    ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">` +
      `<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${refs.svgRid}"/>` +
      `</a:ext></a:extLst>`
    : "";
  const blip = refs.svgRid
    ? (refs.rid ? `<a:blip r:embed="${refs.rid}">${svgExt}</a:blip>` : `<a:blip>${svgExt}</a:blip>`)
    : `<a:blip r:embed="${refs.rid}"/>`;
  const sr = s.srcRect;
  const srcRect = sr && (sr.l || sr.t || sr.r || sr.b)
    ? `<a:srcRect l="${Math.round(sr.l * 100000)}" t="${Math.round(sr.t * 100000)}" r="${Math.round(sr.r * 100000)}" b="${Math.round(sr.b * 100000)}"/>`
    : "";
  const geom = s.geom && s.geom !== "rect" ? s.geom : "rect";
  return `<p:pic><p:nvPicPr><p:cNvPr id="${numId}" name="${esc(s.name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill>${blip}${srcRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(s)}<a:prstGeom prst="${esc(geom)}"><a:avLst>${avLstXml(s.adj)}</a:avLst></a:prstGeom></p:spPr></p:pic>`;
}

function slideXml(slide: SlideModel, imageRels: Map<string, string>, svgRels: Map<string, string>, chartRels: Map<string, string>, picKeys: Map<string, string>): string {
  let bg = "";
  if (slide.background) {
    bg = `<p:bg><p:bgPr>${fillXml(slide.background, imageRels)}<a:effectLst/></p:bgPr></p:bg>`;
  }
  let numId = 2;
  const oneShape = (s: SlideModel["shapes"][number]): string => {
    switch (s.kind) {
      case "pic": {
        const key = picKeys.get(s.id) ?? s.mediaId;
        const rid = imageRels.get(key);
        const svgRid = svgRels.get(key);
        if (!rid && !svgRid) return ""; // media missing — drop rather than emit a dangling r:embed
        return picXml(s, numId++, { rid, svgRid });
      }
      case "table":
        return tableXml(s, numId++);
      case "chart": {
        const rid = chartRels.get(s.id);
        if (!rid) return "";
        return chartFrameXml(s, numId++, rid);
      }
      default:
        return spXml(s, numId++, imageRels);
    }
  };

  // cluster shapes that share a groupId (first-occurrence order) into real p:grpSp
  // elements; the group frame doubles as the child space (chOff/chExt = off/ext),
  // so child coordinates pass through 1:1 with no rewriting
  const clusters: { gid?: string; members: SlideModel["shapes"] }[] = [];
  const byGid = new Map<string, { gid?: string; members: SlideModel["shapes"] }>();
  for (const s of slide.shapes) {
    if (s.groupId) {
      let c = byGid.get(s.groupId);
      if (!c) { c = { gid: s.groupId, members: [] }; byGid.set(s.groupId, c); clusters.push(c); }
      c.members.push(s);
    } else {
      clusters.push({ members: [s] });
    }
  }
  const shapes = clusters.map(c => {
    if (!c.gid || c.members.length < 2) return c.members.map(oneShape).join("");
    const x0 = Math.min(...c.members.map(m => m.x));
    const y0 = Math.min(...c.members.map(m => m.y));
    const x1 = Math.max(...c.members.map(m => m.x + m.w));
    const y1 = Math.max(...c.members.map(m => m.y + m.h));
    const gnum = numId++;
    const xfrm = `<a:xfrm><a:off x="${Math.round(x0)}" y="${Math.round(y0)}"/><a:ext cx="${Math.max(1, Math.round(x1 - x0))}" cy="${Math.max(1, Math.round(y1 - y0))}"/>` +
      `<a:chOff x="${Math.round(x0)}" y="${Math.round(y0)}"/><a:chExt cx="${Math.max(1, Math.round(x1 - x0))}" cy="${Math.max(1, Math.round(y1 - y0))}"/></a:xfrm>`;
    return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${gnum}" name="Group ${gnum}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr>${xfrm}</p:grpSpPr>` +
      c.members.map(oneShape).join("") +
      `</p:grpSp>`;
  }).join("");
  return XML_DECL +
    `<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld>${bg}` +
    `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    shapes +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transitionXml(slide)}</p:sld>`;
}

function transitionXml(slide: SlideModel): string {
  const t = slide.transition;
  if (!t || t.type === "none") return "";
  const inner = t.type === "fade" ? `<p:fade/>` : `<p:push dir="${t.dir ?? "l"}"/>`;
  return `<p:transition spd="${t.speed}">${inner}</p:transition>`;
}

function themeXml(p: Presentation): string {
  const t = p.theme;
  const fmtScheme =
    `<a:fmtScheme name="Office">` +
    `<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>` +
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>` +
    `<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>`;
  return XML_DECL +
    `<a:theme xmlns:a="${NS_A}" name="${esc(t.name)}"><a:themeElements>` +
    `<a:clrScheme name="${esc(t.name)}">` +
    // dk1/lt1 as literal srgbClr (NOT sysClr windowText/window) — sysClr makes
    // PowerPoint substitute the OS black/white and ignore the hex, which flips a
    // custom dark theme to white. (Same form PowerPoint writes for custom themes.)
    `<a:dk1><a:srgbClr val="${t.dk1}"/></a:dk1><a:lt1><a:srgbClr val="${t.lt1}"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="${t.dk2}"/></a:dk2><a:lt2><a:srgbClr val="${t.lt2}"/></a:lt2>` +
    `<a:accent1><a:srgbClr val="${t.accent1}"/></a:accent1><a:accent2><a:srgbClr val="${t.accent2}"/></a:accent2>` +
    `<a:accent3><a:srgbClr val="${t.accent3}"/></a:accent3><a:accent4><a:srgbClr val="${t.accent4}"/></a:accent4>` +
    `<a:accent5><a:srgbClr val="${t.accent5}"/></a:accent5><a:accent6><a:srgbClr val="${t.accent6}"/></a:accent6>` +
    `<a:hlink><a:srgbClr val="${t.hlink}"/></a:hlink><a:folHlink><a:srgbClr val="${t.folHlink}"/></a:folHlink></a:clrScheme>` +
    `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="${esc(t.majorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="${esc(t.minorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
    fmtScheme +
    `</a:themeElements></a:theme>`;
}

function slideMasterXml(): string {
  const lvl = (sz: number) =>
    `<a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>`;
  return XML_DECL +
    `<p:sldMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
    `<p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` +
    `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>` +
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
    `<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:buNone/>${lvl(4400)}</a:lvl1pPr></p:titleStyle>` +
    `<p:bodyStyle><a:lvl1pPr><a:buChar char="•"/>${lvl(2800)}</a:lvl1pPr></p:bodyStyle>` +
    `<p:otherStyle><a:lvl1pPr>${lvl(1800)}</a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;
}

function slideLayoutXml(): string {
  return XML_DECL +
    `<p:sldLayout xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" type="blank" preserve="1">` +
    `<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function rels(items: { id: string; type: string; target: string }[]): string {
  return XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    items.map(i => `<Relationship Id="${i.id}" Type="${i.type}" Target="${i.target}"/>`).join("") +
    `</Relationships>`;
}

function mimeExt(mime: string): string {
  if (mime === "image/jpeg") return "jpeg";
  if (mime === "image/gif") return "gif";
  if (mime === "image/svg+xml") return "svg";
  return "png";
}

export async function buildPptx(pres: Presentation, media: Map<string, MediaItem>): Promise<JSZip> {
  const zip = new JSZip();
  const slideCount = pres.slides.length;

  // ---- media: assign file names ----
  // keys: mediaId, or mediaId#hex for tinted svg graphics (each tint is its own file)
  const usedMedia = new Map<string, { item: MediaItem; file: string; svgFile?: string; svgOnly?: boolean }>();
  let imgNum = 1;
  const addMedia = (key: string, mid: string, tintHex?: string) => {
    if (usedMedia.has(key)) return;
    const item = media.get(mid);
    if (!item) return;
    if (item.mime === "image/svg+xml") {
      const bytes = tintHex
        ? new TextEncoder().encode(tintSvgText(new TextDecoder().decode(item.bytes), tintHex))
        : item.bytes;
      const it = bytes === item.bytes ? item : { ...item, bytes };
      if (item.pngFallback) {
        usedMedia.set(key, { item: it, file: `image${imgNum++}.png`, svgFile: `image${imgNum++}.svg` });
      } else {
        // pure graphic (PowerPoint-style: no bitmap fallback)
        usedMedia.set(key, { item: it, file: `image${imgNum++}.svg`, svgOnly: true });
      }
    } else {
      usedMedia.set(key, { item, file: `image${imgNum++}.${mimeExt(item.mime)}` });
    }
  };
  /** Composite key for a picture: tinted svg graphics get a per-color media file. */
  const picKey = (s: PicShape): { key: string; tintHex?: string } => {
    const item = media.get(s.mediaId);
    if (item && item.mime === "image/svg+xml" && s.svgTint) {
      const hex = cssColorToHex(resolveColor(s.svgTint, pres.theme));
      return { key: `${s.mediaId}#${hex}`, tintHex: hex };
    }
    return { key: s.mediaId };
  };
  for (const slide of pres.slides) {
    if (slide.background?.kind === "image") addMedia(slide.background.mediaId, slide.background.mediaId);
    for (const s of slide.shapes) {
      if (s.kind === "pic") {
        const { key, tintHex } = picKey(s);
        addMedia(key, s.mediaId, tintHex);
      } else if (s.kind === "sp" && s.fill.kind === "image") {
        addMedia(s.fill.mediaId, s.fill.mediaId);
      }
    }
  }

  const chartCount = pres.slides.reduce((n, sl) => n + sl.shapes.filter(s => s.kind === "chart").length, 0);
  const notesCount = pres.slides.filter(s => s.notes?.trim()).length;
  const anyNotes = notesCount > 0;

  // ---- [Content_Types].xml ----
  const overrides: string[] = [
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`,
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`,
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`,
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`,
    `<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>`,
    `<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>`,
    `<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>`,
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`,
  ];
  for (let i = 1; i <= slideCount; i++) {
    overrides.push(`<Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
  }
  for (let i = 1; i <= chartCount; i++) {
    overrides.push(`<Override PartName="/ppt/charts/chart${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
  }
  if (anyNotes) {
    overrides.push(`<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>`);
    overrides.push(`<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`);
    for (let i = 1; i <= notesCount; i++) {
      overrides.push(`<Override PartName="/ppt/notesSlides/notesSlide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`);
    }
  }
  zip.file("[Content_Types].xml", XML_DECL +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Default Extension="jpeg" ContentType="image/jpeg"/>` +
    `<Default Extension="jpg" ContentType="image/jpeg"/>` +
    `<Default Extension="gif" ContentType="image/gif"/>` +
    `<Default Extension="svg" ContentType="image/svg+xml"/>` +
    overrides.join("") + `</Types>`);

  // ---- package rels ----
  zip.file("_rels/.rels", rels([
    { id: "rId1", type: `${RT}/officeDocument`, target: "ppt/presentation.xml" },
    { id: "rId2", type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", target: "docProps/core.xml" },
    { id: "rId3", type: `${RT}/extended-properties`, target: "docProps/app.xml" },
  ]));

  // ---- docProps ----
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  zip.file("docProps/core.xml", XML_DECL +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${esc(pres.title)}</dc:title><dc:creator>Presentation Editor</dc:creator><cp:lastModifiedBy>Presentation Editor</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  zip.file("docProps/app.xml", XML_DECL +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>Presentation Editor</Application><Slides>${slideCount}</Slides><PresentationFormat>Widescreen</PresentationFormat></Properties>`);

  // ---- presentation.xml + rels ----
  const hasNotes = anyNotes;
  const sldIds = pres.slides.map((_, i) =>
    `<p:sldId id="${256 + i}" r:id="rId${2 + i}"/>`).join("");
  zip.file("ppt/presentation.xml", XML_DECL +
    `<p:presentation xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" saveSubsetFonts="1">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    (hasNotes ? `<p:notesMasterIdLst><p:notesMasterId r:id="rId${6 + slideCount}"/></p:notesMasterIdLst>` : "") +
    `<p:sldIdLst>${sldIds}</p:sldIdLst>` +
    `<p:sldSz cx="${pres.slideWidth}" cy="${pres.slideHeight}"/><p:notesSz cx="6858000" cy="9144000"/>` +
    `</p:presentation>`);

  const presRels = [
    { id: "rId1", type: `${RT}/slideMaster`, target: "slideMasters/slideMaster1.xml" },
    ...pres.slides.map((_, i) => ({ id: `rId${2 + i}`, type: `${RT}/slide`, target: `slides/slide${i + 1}.xml` })),
    { id: `rId${2 + slideCount}`, type: `${RT}/presProps`, target: "presProps.xml" },
    { id: `rId${3 + slideCount}`, type: `${RT}/viewProps`, target: "viewProps.xml" },
    { id: `rId${4 + slideCount}`, type: `${RT}/theme`, target: "theme/theme1.xml" },
    { id: `rId${5 + slideCount}`, type: `${RT}/tableStyles`, target: "tableStyles.xml" },
    ...(hasNotes ? [{ id: `rId${6 + slideCount}`, type: `${RT}/notesMaster`, target: "notesMasters/notesMaster1.xml" }] : []),
  ];
  zip.file("ppt/_rels/presentation.xml.rels", rels(presRels));

  // ---- static props parts ----
  zip.file("ppt/presProps.xml", XML_DECL + `<p:presentationPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"/>`);
  zip.file("ppt/viewProps.xml", XML_DECL + `<p:viewPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`);
  zip.file("ppt/tableStyles.xml", XML_DECL + `<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);

  // ---- theme / master / layout ----
  zip.file("ppt/theme/theme1.xml", themeXml(pres));
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMasterXml());
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", rels([
    { id: "rId1", type: `${RT}/slideLayout`, target: "../slideLayouts/slideLayout1.xml" },
    { id: "rId2", type: `${RT}/theme`, target: "../theme/theme1.xml" },
  ]));
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayoutXml());
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", rels([
    { id: "rId1", type: `${RT}/slideMaster`, target: "../slideMasters/slideMaster1.xml" },
  ]));

  // ---- media files ----
  for (const { item, file, svgFile } of usedMedia.values()) {
    if (svgFile) {
      zip.file(`ppt/media/${file}`, item.pngFallback!); // bitmap fallback blip
      zip.file(`ppt/media/${svgFile}`, item.bytes);      // the vector itself
    } else {
      zip.file(`ppt/media/${file}`, item.bytes);
    }
  }

  // ---- notes master (only when needed) ----
  if (hasNotes) {
    zip.file("ppt/notesMasters/notesMaster1.xml", XML_DECL +
      `<p:notesMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
      `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `</p:notesMaster>`);
    zip.file("ppt/notesMasters/_rels/notesMaster1.xml.rels", rels([
      { id: "rId1", type: `${RT}/theme`, target: "../theme/theme2.xml" },
    ]));
    zip.file("ppt/theme/theme2.xml", themeXml(pres));
  }

  // ---- slides (+ chart parts, notes slides) ----
  let chartNum = 1;
  let notesNum = 1;
  pres.slides.forEach((slide, i) => {
    const slideRels = [
      { id: "rId1", type: `${RT}/slideLayout`, target: "../slideLayouts/slideLayout1.xml" },
    ];
    const imageRels = new Map<string, string>();
    const svgRels = new Map<string, string>();
    const chartRels = new Map<string, string>();
    const picKeys = new Map<string, string>(); // shape id -> media key
    let nextR = 2;
    // every image used on this slide: pictures, shape picture-fills, background
    const slideMediaKeys = new Set<string>();
    if (slide.background?.kind === "image") slideMediaKeys.add(slide.background.mediaId);
    for (const s of slide.shapes) {
      if (s.kind === "pic") {
        const { key } = picKey(s);
        picKeys.set(s.id, key);
        slideMediaKeys.add(key);
      } else if (s.kind === "sp" && s.fill.kind === "image") {
        slideMediaKeys.add(s.fill.mediaId);
      }
    }
    for (const key of slideMediaKeys) {
      const um = usedMedia.get(key);
      if (!um) continue;
      if (um.svgOnly) {
        // pure graphic: single svg rel, referenced from the asvg extension only
        const srid = `rId${nextR++}`;
        svgRels.set(key, srid);
        slideRels.push({ id: srid, type: `${RT}/image`, target: `../media/${um.file}` });
        continue;
      }
      const rid = `rId${nextR++}`;
      imageRels.set(key, rid);
      slideRels.push({ id: rid, type: `${RT}/image`, target: `../media/${um.file}` });
      if (um.svgFile) {
        const srid = `rId${nextR++}`;
        svgRels.set(key, srid);
        slideRels.push({ id: srid, type: `${RT}/image`, target: `../media/${um.svgFile}` });
      }
    }
    for (const s of slide.shapes) {
      if (s.kind === "chart") {
        const rid = `rId${nextR++}`;
        chartRels.set(s.id, rid);
        const file = `chart${chartNum++}.xml`;
        zip.file(`ppt/charts/${file}`, chartSpaceXml(s));
        slideRels.push({ id: rid, type: `${RT}/chart`, target: `../charts/${file}` });
      }
    }
    if (slide.notes?.trim()) {
      const nfile = `notesSlide${notesNum++}.xml`;
      const paras = slide.notes.split(/\r?\n/).map(line =>
        `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${esc(line)}</a:t></a:r></a:p>`).join("");
      zip.file(`ppt/notesSlides/${nfile}`, XML_DECL +
        `<p:notes xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
        `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>` +
        `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`);
      zip.file(`ppt/notesSlides/_rels/${nfile}.rels`, rels([
        { id: "rId1", type: `${RT}/notesMaster`, target: "../notesMasters/notesMaster1.xml" },
        { id: "rId2", type: `${RT}/slide`, target: `../slides/slide${i + 1}.xml` },
      ]));
      slideRels.push({ id: `rId${nextR++}`, type: `${RT}/notesSlide`, target: `../notesSlides/${nfile}` });
    }
    zip.file(`ppt/slides/slide${i + 1}.xml`, slideXml(slide, imageRels, svgRels, chartRels, picKeys));
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels(slideRels));
  });

  return zip;
}

export async function exportPptxBlob(pres: Presentation, media: Map<string, MediaItem>): Promise<Blob> {
  const zip = await buildPptx(pres, media);
  return zip.generateAsync({ type: "blob", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
}
