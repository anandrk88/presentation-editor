import JSZip from "jszip";
import type {
  ChartKind, ChartSeries, ChartShape, ColorRef, ColorTheme, Fill, LineProps,
  MediaItem, Paragraph, Presentation, PresetGeom, Run, SchemeSlot, SlideModel,
  SpShape, TableCell, TableShape, TextBody,
} from "../model/types";
import { OFFICE_THEME, nextId } from "../model/defaults";
import { hasPreset } from "../render/presetGeom";

/**
 * OOXML PresentationML reader: pptx (zip) -> document model.
 * Follows package relationships the same way OnlyOffice's loader does:
 * /_rels/.rels -> presentation.xml -> slide parts; placeholder shapes without
 * an xfrm resolve their geometry through slideLayout -> slideMaster.
 */

type DOMParserCtor = new () => { parseFromString(s: string, t: string): Document };
let DOMParserImpl: DOMParserCtor | undefined =
  typeof DOMParser !== "undefined" ? (DOMParser as unknown as DOMParserCtor) : undefined;

type XMLSerializerCtor = new () => { serializeToString(n: Node): string };
let XMLSerializerImpl: XMLSerializerCtor | undefined =
  typeof XMLSerializer !== "undefined" ? (XMLSerializer as unknown as XMLSerializerCtor) : undefined;

/** Allows Node-based tests to inject @xmldom/xmldom. */
export function setDOMParser(ctor: DOMParserCtor) {
  DOMParserImpl = ctor;
}
export function setXMLSerializer(ctor: XMLSerializerCtor) {
  XMLSerializerImpl = ctor;
}

function parseXml(s: string): Document {
  if (!DOMParserImpl) throw new Error("No DOMParser available");
  return new DOMParserImpl().parseFromString(s, "application/xml") as unknown as Document;
}

// ---------- namespace-agnostic element helpers ----------
function kids(el: Element | null | undefined, local: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const n = el.childNodes[i] as Element;
    if (n.nodeType === 1 && (n.localName === local)) out.push(n);
  }
  return out;
}
function kid(el: Element | null | undefined, ...path: string[]): Element | null {
  let cur: Element | null = el ?? null;
  for (const p of path) {
    if (!cur) return null;
    cur = kids(cur, p)[0] ?? null;
  }
  return cur;
}
function attr(el: Element | null | undefined, name: string): string | null {
  return el ? el.getAttribute(name) : null;
}
function iattr(el: Element | null | undefined, name: string, dflt = 0): number {
  const v = attr(el, name);
  return v === null ? dflt : parseInt(v, 10);
}

// ---------- base64 (browser + node) ----------
function bytesToDataUrl(bytes: Uint8Array, mime: string): string {
  let b64: string;
  if (typeof btoa === "function") {
    let bin = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)));
    }
    b64 = btoa(bin);
  } else {
    // node
    b64 = (globalThis as any).Buffer.from(bytes).toString("base64");
  }
  return `data:${mime};base64,${b64}`;
}

function extToMime(path: string): string {
  const ext = path.split(".").pop()?.toLowerCase();
  switch (ext) {
    case "jpg": case "jpeg": return "image/jpeg";
    case "gif": return "image/gif";
    case "svg": return "image/svg+xml";
    case "bmp": return "image/bmp";
    case "tiff": case "tif": return "image/tiff";
    default: return "image/png";
  }
}

// ---------- relationships ----------
function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir ? baseDir.split("/") : []).concat(target.split("/"));
  const out: string[] = [];
  for (const p of parts) {
    if (p === "." || p === "") continue;
    if (p === "..") out.pop();
    else out.push(p);
  }
  return out.join("/");
}

interface Rel { id: string; type: string; target: string }

async function parseRels(zip: JSZip, partPath: string): Promise<Map<string, Rel>> {
  const dir = partPath.includes("/") ? partPath.slice(0, partPath.lastIndexOf("/")) : "";
  const relPath = `${dir ? dir + "/" : ""}_rels/${partPath.split("/").pop()}.rels`;
  const map = new Map<string, Rel>();
  const file = zip.file(relPath);
  if (!file) return map;
  const doc = parseXml(await file.async("text"));
  const root = doc.documentElement;
  for (const r of kids(root, "Relationship")) {
    const id = attr(r, "Id")!;
    const type = attr(r, "Type") ?? "";
    const mode = attr(r, "TargetMode");
    let target = attr(r, "Target") ?? "";
    if (mode !== "External") target = resolveTarget(dir, target);
    map.set(id, { id, type, target });
  }
  return map;
}

// ---------- colors / fills ----------
const SCHEME_ALIAS: Record<string, SchemeSlot | undefined> = {
  tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2",
  dk1: "dk1", lt1: "lt1", dk2: "dk2", lt2: "lt2",
  accent1: "accent1", accent2: "accent2", accent3: "accent3",
  accent4: "accent4", accent5: "accent5", accent6: "accent6",
  hlink: "hlink", folHlink: "folHlink",
};

const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const PRESET_CLR: Record<string, string> = {
  white: "FFFFFF", black: "000000", red: "FF0000", green: "008000", blue: "0000FF",
  yellow: "FFFF00", gray: "808080", grey: "808080", orange: "FFA500", purple: "800080",
};

function parseColorChoice(parent: Element | null): ColorRef | null {
  if (!parent) return null;
  for (let i = 0; i < parent.childNodes.length; i++) {
    const n = parent.childNodes[i] as Element;
    if (n.nodeType !== 1) continue;
    const lumMod = kid(n, "lumMod") ? iattr(kid(n, "lumMod"), "val") / 1000 : undefined;
    const lumOff = kid(n, "lumOff") ? iattr(kid(n, "lumOff"), "val") / 1000 : undefined;
    const alpha = kid(n, "alpha") ? iattr(kid(n, "alpha"), "val") / 1000 : undefined;
    switch (n.localName) {
      case "srgbClr":
        return { kind: "srgb", hex: (attr(n, "val") ?? "000000").toUpperCase(), lumMod, lumOff, alpha };
      case "schemeClr": {
        const raw = attr(n, "val") ?? "accent1";
        if (raw === "phClr") return { kind: "scheme", slot: "accent1", lumMod, lumOff, alpha };
        const slot = SCHEME_ALIAS[raw];
        return slot ? { kind: "scheme", slot, lumMod, lumOff, alpha } : { kind: "srgb", hex: "808080", alpha };
      }
      case "sysClr":
        return { kind: "srgb", hex: (attr(n, "lastClr") ?? "000000").toUpperCase(), alpha };
      case "prstClr":
        return { kind: "srgb", hex: PRESET_CLR[attr(n, "val") ?? ""] ?? "808080", alpha };
    }
  }
  return null;
}

/** Parse fill from an spPr-like element (first matching fill child). */
function parseFill(el: Element | null, groupFill?: Fill): Fill | undefined {
  if (!el) return undefined;
  for (let i = 0; i < el.childNodes.length; i++) {
    const n = el.childNodes[i] as Element;
    if (n.nodeType !== 1) continue;
    switch (n.localName) {
      case "noFill": return { kind: "none" };
      case "solidFill": {
        const c = parseColorChoice(n);
        return c ? { kind: "solid", color: c } : { kind: "none" };
      }
      case "gradFill": {
        const gs = kid(n, "gsLst");
        const stops = (gs ? kids(gs, "gs") : [])
          .map(g => ({ pos: iattr(g, "pos", 0) / 1000, color: parseColorChoice(g) }))
          .filter((s): s is { pos: number; color: ColorRef } => s.color !== null)
          .sort((a, b) => a.pos - b.pos);
        if (!stops.length) return undefined;
        const lin = kid(n, "lin");
        const angle = lin ? iattr(lin, "ang", 0) / 60000 : 90; // a:path (radial) approximated as vertical linear
        return { kind: "gradient", stops, angle };
      }
      case "blipFill": {
        // mediaId carries the r:embed id here — the slide loop swaps it for a real media id
        const blip = kid(n, "blip");
        const embed = blip?.getAttributeNS?.(NS_REL, "embed") ?? attr(blip, "r:embed");
        if (embed) return { kind: "image", mediaId: embed, tile: kid(n, "tile") ? true : undefined };
        return { kind: "solid", color: { kind: "srgb", hex: "D9D9D9" } };
      }
      case "grpFill": return groupFill; // inherit the enclosing group's fill
      case "pattFill": {
        const fg = parseColorChoice(kid(n, "fgClr")) ?? { kind: "srgb" as const, hex: "808080" };
        const bg = parseColorChoice(kid(n, "bgClr")) ?? { kind: "srgb" as const, hex: "FFFFFF" };
        return { kind: "pattern", prst: attr(n, "prst") ?? "ltDnDiag", fg, bg };
      }
    }
  }
  return undefined;
}

/** a:custGeom -> SVG path data in the path's own coordinate space. */
function parseCustGeom(custGeom: Element): { w: number; h: number; d: string } | null {
  const pathLst = kid(custGeom, "pathLst");
  if (!pathLst) return null;
  const deg = (v: number) => (v / 60000) * (Math.PI / 180);
  let w = 0, h = 0;
  let d = "";
  for (const path of kids(pathLst, "path")) {
    w = Math.max(w, iattr(path, "w", 0));
    h = Math.max(h, iattr(path, "h", 0));
    let cx = 0, cy = 0;
    const pt = (el: Element | null): [number, number] =>
      [iattr(el, "x", 0), iattr(el, "y", 0)];
    for (let i = 0; i < path.childNodes.length; i++) {
      const c = path.childNodes[i] as Element;
      if (c.nodeType !== 1) continue;
      switch (c.localName) {
        case "moveTo": { [cx, cy] = pt(kid(c, "pt")); d += `M${cx} ${cy}`; break; }
        case "lnTo": { [cx, cy] = pt(kid(c, "pt")); d += `L${cx} ${cy}`; break; }
        case "cubicBezTo": {
          const ps = kids(c, "pt").map(p => pt(p));
          if (ps.length === 3) {
            d += `C${ps[0][0]} ${ps[0][1]} ${ps[1][0]} ${ps[1][1]} ${ps[2][0]} ${ps[2][1]}`;
            [cx, cy] = ps[2];
          }
          break;
        }
        case "quadBezTo": {
          const ps = kids(c, "pt").map(p => pt(p));
          if (ps.length === 2) {
            d += `Q${ps[0][0]} ${ps[0][1]} ${ps[1][0]} ${ps[1][1]}`;
            [cx, cy] = ps[1];
          }
          break;
        }
        case "arcTo": {
          const wR = iattr(c, "wR", 0), hR = iattr(c, "hR", 0);
          const st = deg(iattr(c, "stAng", 0)), sw = deg(iattr(c, "swAng", 0));
          const ecx = cx - wR * Math.cos(st), ecy = cy - hR * Math.sin(st);
          const ex = ecx + wR * Math.cos(st + sw), ey = ecy + hR * Math.sin(st + sw);
          const large = Math.abs(sw) > Math.PI ? 1 : 0;
          const sweep = sw > 0 ? 1 : 0;
          d += `A${wR} ${hR} 0 ${large} ${sweep} ${ex.toFixed(2)} ${ey.toFixed(2)}`;
          cx = ex; cy = ey;
          break;
        }
        case "close": d += "Z"; break;
      }
    }
  }
  if (!d) return null;
  return { w: w || 1, h: h || 1, d };
}

function parseLn(spPr: Element | null): LineProps | undefined {
  const ln = kid(spPr, "ln");
  if (!ln) return undefined;
  const widthPt = iattr(ln, "w", 9525) / 12700;
  const fill = parseFill(ln) ?? { kind: "none" as const };
  const dashEl = kid(ln, "prstDash");
  const dashVal = attr(dashEl, "val");
  const dash = dashVal && dashVal.includes("dash") ? "dash" as const : dashVal === "sysDot" || dashVal === "dot" ? "dot" as const : "solid" as const;
  return { fill, widthPt: Math.max(0.25, widthPt), dash };
}

// ---------- geometry ----------
/** avLst overrides: <a:gd name="adj" fmla="val 16667"/> */
function parseAdj(prstGeom: Element | null): Record<string, number> | undefined {
  const avLst = kid(prstGeom, "avLst");
  if (!avLst) return undefined;
  const out: Record<string, number> = {};
  for (const g of kids(avLst, "gd")) {
    const name = attr(g, "name");
    const m = /^val\s+(-?[\d.]+)$/.exec(attr(g, "fmla") ?? "");
    // guard against prototype-pollution keys from a crafted file
    if (name && m && name !== "__proto__" && name !== "constructor" && name !== "prototype") {
      out[name] = parseFloat(m[1]);
    }
  }
  return Object.keys(out).length ? out : undefined;
}

// ---------- text ----------
interface TextDefaults { sizePt: number; font: string; color: ColorRef; align?: Paragraph["align"]; bullet?: boolean; bold?: boolean; italic?: boolean }

/** One inheritance layer of default run properties (a:defRPr). */
interface LevelDefault {
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  color?: ColorRef;
  font?: string;
}
/** Per indent level (0-8), from a:lstStyle / p:txStyles. */
type LstDefaults = (LevelDefault | undefined)[];

function parseDefRPrEl(el: Element | null): LevelDefault | undefined {
  if (!el) return undefined;
  const out: LevelDefault = {};
  const sz = iattr(el, "sz", 0);
  if (sz) out.sizePt = sz / 100;
  const b = attr(el, "b");
  if (b === "1") out.bold = true; else if (b === "0") out.bold = false;
  const i = attr(el, "i");
  if (i === "1") out.italic = true; else if (i === "0") out.italic = false;
  const fill = kid(el, "solidFill");
  const c = fill ? parseColorChoice(fill) : null;
  if (c) out.color = c;
  const latin = attr(kid(el, "latin"), "typeface");
  if (latin) out.font = latin;
  return Object.keys(out).length ? out : undefined;
}

function parseLstStyle(lst: Element | null): LstDefaults | undefined {
  if (!lst) return undefined;
  const out: LstDefaults = [];
  let any = false;
  for (let lvl = 0; lvl < 9; lvl++) {
    const p = kid(lst, `lvl${lvl + 1}pPr`);
    const d = parseDefRPrEl(kid(p, "defRPr"));
    if (d) { out[lvl] = d; any = true; }
  }
  return any ? out : undefined;
}

/** Later layers override earlier ones, per level. */
function mergeLst(...layers: (LstDefaults | undefined)[]): LstDefaults | undefined {
  const out: LstDefaults = [];
  let any = false;
  for (let lvl = 0; lvl < 9; lvl++) {
    const merged: LevelDefault = {};
    for (const layer of layers) {
      if (layer?.[lvl]) Object.assign(merged, layer[lvl]);
    }
    if (Object.keys(merged).length) { out[lvl] = merged; any = true; }
  }
  return any ? out : undefined;
}

function parseRunProps(rPr: Element | null, dflt: TextDefaults, theme: ThemeFonts): Omit<Run, "text"> {
  const sz = rPr ? iattr(rPr, "sz", 0) : 0;
  const latin = kid(rPr, "latin");
  // +mj-lt / +mn-lt stay symbolic — they re-font when the theme fonts change
  const font = attr(latin, "typeface") ?? dflt.font;
  void theme;
  const colorEl = kid(rPr, "solidFill");
  const color = (colorEl && parseColorChoice(colorEl)) || dflt.color;
  const uVal = attr(rPr, "u");
  const strikeVal = attr(rPr, "strike");
  const baselineRaw = rPr ? iattr(rPr, "baseline", 0) : 0;
  const highlightEl = kid(rPr, "highlight");
  const bAttr = attr(rPr, "b");
  const iAttr = attr(rPr, "i");
  return {
    sizePt: sz ? sz / 100 : dflt.sizePt,
    font,
    color,
    // explicit attribute wins; otherwise inherit from the defRPr chain
    bold: bAttr !== null ? (bAttr === "1" || undefined) : (dflt.bold || undefined),
    italic: iAttr !== null ? (iAttr === "1" || undefined) : (dflt.italic || undefined),
    underline: (uVal && uVal !== "none") || undefined,
    strike: (strikeVal && strikeVal !== "noStrike") || undefined,
    baseline: baselineRaw ? baselineRaw / 1000 : undefined,
    highlight: highlightEl ? parseColorChoice(highlightEl) ?? undefined : undefined,
  };
}

function parseTxBody(txBody: Element | null, dflt: TextDefaults, theme: ThemeFonts, inheritedLst?: LstDefaults): TextBody | undefined {
  if (!txBody) return undefined;
  const bodyPr = kid(txBody, "bodyPr");
  const anchorRaw = attr(bodyPr, "anchor");
  const anchor = anchorRaw === "ctr" ? "ctr" : anchorRaw === "b" ? "b" : "t";
  const wrap = attr(bodyPr, "wrap") !== "none";
  const insets: [number, number, number, number] = [
    iattr(bodyPr, "lIns", 91440), iattr(bodyPr, "tIns", 45720),
    iattr(bodyPr, "rIns", 91440), iattr(bodyPr, "bIns", 45720),
  ];
  const columns = iattr(bodyPr, "numCol", 1);
  const colSpacing = iattr(bodyPr, "spcCol", 0);
  // size/style inheritance: placeholder chain (layout/master) <- shape's own lstStyle
  const lst = mergeLst(inheritedLst, parseLstStyle(kid(txBody, "lstStyle")));
  const paragraphs: Paragraph[] = [];
  for (const p of kids(txBody, "p")) {
    const pPr = kid(p, "pPr");
    const algnRaw = attr(pPr, "algn");
    const align: Paragraph["align"] =
      algnRaw === "ctr" ? "ctr" : algnRaw === "r" ? "r" : algnRaw === "just" ? "just" : (dflt.align ?? "l");
    const level = iattr(pPr, "lvl", 0);
    let bullet: Paragraph["bullet"] = dflt.bullet ? "char" : "none";
    if (pPr) {
      if (kid(pPr, "buNone")) bullet = "none";
      else if (kid(pPr, "buChar")) bullet = "char";
      else if (kid(pPr, "buAutoNum")) bullet = "num";
    }
    const lnSpcPct = kid(pPr, "lnSpc", "spcPct");
    const lineSpacingPct = lnSpcPct ? iattr(lnSpcPct, "val") / 1000 : undefined;

    // effective defaults for this paragraph: base <- lstStyle[level] <- pPr/defRPr
    const lvlDef = lst?.[Math.min(level, 8)];
    const pDef = parseDefRPrEl(kid(pPr, "defRPr"));
    const eff: TextDefaults = {
      sizePt: pDef?.sizePt ?? lvlDef?.sizePt ?? dflt.sizePt,
      font: pDef?.font ?? lvlDef?.font ?? dflt.font,
      color: pDef?.color ?? lvlDef?.color ?? dflt.color,
      align: dflt.align,
      bullet: dflt.bullet,
      bold: pDef?.bold ?? lvlDef?.bold ?? dflt.bold,
      italic: pDef?.italic ?? lvlDef?.italic ?? dflt.italic,
    };

    const runs: Run[] = [];
    for (let i = 0; i < p.childNodes.length; i++) {
      const n = p.childNodes[i] as Element;
      if (n.nodeType !== 1) continue;
      if (n.localName === "r") {
        const t = kid(n, "t");
        const text = t?.textContent ?? "";
        runs.push({ text, ...parseRunProps(kid(n, "rPr"), eff, theme) });
      } else if (n.localName === "br") {
        if (runs.length) runs[runs.length - 1] = { ...runs[runs.length - 1], text: runs[runs.length - 1].text + "\n" };
        else runs.push({ text: "\n", ...parseRunProps(null, eff, theme) });
      } else if (n.localName === "fld") {
        const t = kid(n, "t");
        runs.push({ text: t?.textContent ?? "", ...parseRunProps(kid(n, "rPr"), eff, theme) });
      }
    }
    // empty paragraph: keep a style carrier so the caret/typing inherits the right size
    if (!runs.length) {
      const endPr = parseDefRPrEl(kid(p, "endParaRPr"));
      runs.push({
        text: "",
        sizePt: endPr?.sizePt ?? eff.sizePt,
        font: endPr?.font ?? eff.font,
        color: endPr?.color ?? eff.color,
        bold: (endPr?.bold ?? eff.bold) || undefined,
        italic: (endPr?.italic ?? eff.italic) || undefined,
      });
    }
    paragraphs.push({ runs, align, bullet, level, lineSpacingPct });
  }
  if (!paragraphs.length) paragraphs.push({ runs: [], align: "l", bullet: "none", level: 0 });
  return {
    paragraphs, anchor, wrap, insets,
    columns: columns > 1 ? columns : undefined,
    colSpacing: columns > 1 && colSpacing ? colSpacing : undefined,
  };
}

// ---------- tables ----------
function parseTableEl(tbl: Element, dflt: TextDefaults, theme: ThemeFonts): Pick<TableShape, "colW" | "rowH" | "cells" | "firstRow" | "bandRow" | "totalRow" | "firstCol" | "lastCol" | "bandCol"> {
  const tblPr = kid(tbl, "tblPr");
  const firstRow = attr(tblPr, "firstRow") === "1";
  const bandRow = attr(tblPr, "bandRow") === "1";
  const totalRow = attr(tblPr, "lastRow") === "1" || undefined;
  const firstCol = attr(tblPr, "firstCol") === "1" || undefined;
  const lastCol = attr(tblPr, "lastCol") === "1" || undefined;
  const bandCol = attr(tblPr, "bandCol") === "1" || undefined;
  const colW = kids(kid(tbl, "tblGrid"), "gridCol").map(c => iattr(c, "w", 914400));
  const rowH: number[] = [];
  const cells: TableCell[][] = [];
  for (const tr of kids(tbl, "tr")) {
    rowH.push(iattr(tr, "h", 370840));
    const row: TableCell[] = [];
    for (const tc of kids(tr, "tc")) {
      const text = parseTxBody(kid(tc, "txBody"), dflt, theme) ?? {
        paragraphs: [{ runs: [], align: "l" as const, bullet: "none" as const, level: 0 }],
        anchor: "ctr" as const, wrap: true, insets: [91440, 45720, 91440, 45720] as [number, number, number, number],
      };
      const tcPr = kid(tc, "tcPr");
      const anchorRaw = attr(tcPr, "anchor");
      if (anchorRaw === "ctr" || anchorRaw === "b") text.anchor = anchorRaw;
      const merged = attr(tc, "hMerge") === "1" ? "h" as const : attr(tc, "vMerge") === "1" ? "v" as const : undefined;
      row.push({
        text,
        fill: parseFill(tcPr),
        gridSpan: iattr(tc, "gridSpan", 1) > 1 ? iattr(tc, "gridSpan", 1) : undefined,
        rowSpan: iattr(tc, "rowSpan", 1) > 1 ? iattr(tc, "rowSpan", 1) : undefined,
        merged,
      });
    }
    cells.push(row);
  }
  // normalize ragged rows
  const nCols = Math.max(colW.length, ...cells.map(r => r.length), 1);
  while (colW.length < nCols) colW.push(914400);
  for (const row of cells) {
    while (row.length < nCols) row.push({ text: { paragraphs: [{ runs: [], align: "l", bullet: "none", level: 0 }], anchor: "ctr", wrap: true, insets: [91440, 45720, 91440, 45720] } });
  }
  return { colW, rowH, cells, firstRow, bandRow, totalRow, firstCol, lastCol, bandCol };
}

// ---------- charts ----------
/** Collect pt values (idx -> text) under the first strRef/strCache/strLit/numRef/numCache/numLit descendant. */
function chartPts(node: Element | null): string[] {
  if (!node) return [];
  const cacheTags = ["strCache", "strLit", "numCache", "numLit"];
  let holder: Element | null = null;
  const walk = (el: Element) => {
    if (holder) return;
    if (cacheTags.includes(el.localName)) { holder = el; return; }
    for (let i = 0; i < el.childNodes.length; i++) {
      const c = el.childNodes[i] as Element;
      if (c.nodeType === 1) walk(c);
    }
  };
  walk(node);
  if (!holder) return [];
  const out: string[] = [];
  for (const pt of kids(holder, "pt")) {
    const idx = iattr(pt, "idx", out.length);
    out[idx] = kid(pt, "v")?.textContent ?? "";
  }
  for (let i = 0; i < out.length; i++) if (out[i] === undefined) out[i] = "";
  return out;
}

const CHART_TYPE_ELS: [string, ChartKind][] = [
  ["barChart", "column"], ["lineChart", "line"], ["areaChart", "area"],
  ["pieChart", "pie"], ["doughnutChart", "doughnut"],
  ["scatterChart", "scatter"], ["radarChart", "radar"],
];

/** Collect concatenated a:t text under an element (chart/axis titles). */
function richText(el: Element | null): string | undefined {
  if (!el) return undefined;
  const texts: string[] = [];
  const walkT = (e: Element) => {
    for (let i = 0; i < e.childNodes.length; i++) {
      const c = e.childNodes[i] as Element;
      if (c.nodeType !== 1) continue;
      if (c.localName === "t") texts.push(c.textContent ?? "");
      else walkT(c);
    }
  };
  walkT(el);
  const out = texts.join("");
  return out || undefined;
}

function parseChartDoc(doc: Document): Pick<ChartShape, "chart" | "categories" | "series" | "title" | "legend" | "grouping" | "marker" | "smooth" | "radarStyle" | "labelSizePt" | "legendPos" | "dataLabels" | "errorBarsPct" | "axisTitleX" | "axisTitleY" | "hideAxisX" | "hideAxisY" | "chartFill" | "chartBorder" | "plotFill" | "plotBorder" | "markerSizePt" | "pointColors" | "gridColor" | "hideGridlines"> | null {
  const chart = kid(doc.documentElement, "chart");
  const plotArea = kid(chart, "plotArea");
  if (!plotArea) return null;
  let kind: ChartKind | null = null;
  let typeEl: Element | null = null;
  for (const [tag, k] of CHART_TYPE_ELS) {
    const el = kid(plotArea, tag);
    if (el) {
      typeEl = el;
      kind = tag === "barChart" && attr(kid(el, "barDir"), "val") === "bar" ? "bar" : k;
      break;
    }
  }
  if (!typeEl || !kind) return null;

  const isScatter = kind === "scatter";
  const series: ChartSeries[] = [];
  let categories: string[] = [];
  let anyMarker = false;
  let anySmooth = false;
  let errorBarsPct: number | undefined;
  for (const ser of kids(typeEl, "ser")) {
    const tx = kid(ser, "tx");
    const name = chartPts(tx)[0] ?? kid(tx, "v")?.textContent ?? `Series ${series.length + 1}`;
    const cats = chartPts(kid(ser, isScatter ? "xVal" : "cat"));
    if (cats.length > categories.length) categories = cats;
    const values = chartPts(kid(ser, isScatter ? "yVal" : "val")).map(v => parseFloat(v) || 0);
    const colorEl = kid(ser, "spPr", "solidFill");
    const sym = attr(kid(ser, "marker", "symbol"), "val");
    if (sym && sym !== "none") anyMarker = true;
    if (attr(kid(ser, "smooth"), "val") === "1") anySmooth = true;
    const eb = kid(ser, "errBars");
    if (eb && attr(kid(eb, "errValType"), "val") === "percentage") {
      const v = parseFloat(attr(kid(eb, "val"), "val") ?? "");
      if (Number.isFinite(v) && v > 0) errorBarsPct = v;
    }
    // series line style (line/scatter/radar): a:ln width + dash
    const serLn = kid(ser, "spPr", "ln");
    const lineWidthPt = serLn && attr(serLn, "w") ? iattr(serLn, "w") / 12700 : undefined;
    const serDashVal = attr(kid(serLn, "prstDash"), "val");
    const serDash = serDashVal && serDashVal.includes("dash") ? "dash" as const
      : serDashVal === "dot" || serDashVal === "sysDot" ? "dot" as const : undefined;
    series.push({
      name, values,
      color: colorEl ? parseColorChoice(colorEl) ?? undefined : undefined,
      lineWidthPt: lineWidthPt && Math.abs(lineWidthPt - 2.25) > 0.01 ? Math.round(lineWidthPt * 100) / 100 : undefined,
      dash: serDash,
    });
  }
  // marker size (first series carrying one)
  let markerSizePt: number | undefined;
  for (const ser of kids(typeEl, "ser")) {
    const sz = iattr(kid(ser, "marker", "size"), "val", 0);
    if (sz && sz !== 5) { markerSizePt = sz; break; }
  }

  // data labels (group-level)
  const dLblsEl = kid(typeEl, "dLbls");
  const dataLabels = dLblsEl
    ? attr(kid(dLblsEl, "showVal"), "val") === "1" || attr(kid(dLblsEl, "showPercent"), "val") === "1" || undefined
    : undefined;

  // axes: read titles + deletion by position (b = horizontal, l = vertical)
  let axisTitleX: string | undefined, axisTitleY: string | undefined;
  let hideAxisX: boolean | undefined, hideAxisY: boolean | undefined;
  for (const axTag of ["catAx", "valAx", "dateAx"]) {
    for (const ax of kids(plotArea, axTag)) {
      const pos = attr(kid(ax, "axPos"), "val");
      const titleText = richText(kid(ax, "title"));
      const deleted = attr(kid(ax, "delete"), "val") === "1";
      if (pos === "b" || pos === "t") {
        if (titleText) axisTitleX = titleText;
        if (deleted) hideAxisX = true;
      } else if (pos === "l" || pos === "r") {
        if (titleText) axisTitleY = titleText;
        if (deleted) hideAxisY = true;
      }
    }
  }
  // bar charts swap: the left axis is the category axis (our Y), bottom is values (our X)
  // — handled naturally because we keyed off axPos, not axis type.
  if (!categories.length) categories = series[0]?.values.map((_, i) => `Category ${i + 1}`) ?? [];

  // gridlines: any axis carrying c:majorGridlines (+ optional spPr/ln color);
  // a cartesian/radar chart with none anywhere has them turned off
  let gridColor: ColorRef | undefined;
  let sawGridlines = false;
  for (const axTag of ["catAx", "valAx", "dateAx"]) {
    for (const ax of kids(plotArea, axTag)) {
      const mg = kid(ax, "majorGridlines");
      if (!mg) continue;
      sawGridlines = true;
      const sf = kid(mg, "spPr", "ln", "solidFill");
      if (sf) gridColor = parseColorChoice(sf) ?? gridColor;
    }
  }
  const hideGridlines = kind !== "pie" && kind !== "doughnut" && !sawGridlines ? true : undefined;

  // pie/doughnut slice colors: c:dPt overrides on the first series.
  // Colors matching the automatic accent cycle stay null (automatic).
  let pointColors: (ColorRef | null)[] | undefined;
  if (kind === "pie" || kind === "doughnut") {
    const accents: SchemeSlot[] = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    const arr: (ColorRef | null)[] = categories.map(() => null);
    let anyExplicit = false;
    for (const dp of kids(kid(typeEl, "ser"), "dPt")) {
      const idx = iattr(kid(dp, "idx"), "val", -1);
      const sf = kid(dp, "spPr", "solidFill");
      if (idx < 0 || !sf) continue;
      const c = parseColorChoice(sf);
      if (!c) continue;
      const auto = c.kind === "scheme" && c.slot === accents[idx % accents.length] && !c.lumMod && !c.lumOff && !c.alpha;
      while (arr.length <= idx) arr.push(null);
      arr[idx] = auto ? null : c;
      if (!auto) anyExplicit = true;
    }
    if (anyExplicit) pointColors = arr;
  }

  // variant flags
  const groupingRaw = attr(kid(typeEl, "grouping"), "val");
  const grouping = groupingRaw === "stacked" || groupingRaw === "percentStacked" ? groupingRaw : undefined;
  let marker: boolean | undefined;
  let smooth: boolean | undefined;
  let radarStyle: ChartShape["radarStyle"];
  if (kind === "line") {
    marker = anyMarker || undefined;
    smooth = anySmooth || undefined;
  } else if (kind === "scatter") {
    marker = anyMarker || undefined;
    const style = attr(kid(typeEl, "scatterStyle"), "val") ?? "marker";
    smooth = style.startsWith("smooth") ? true : style.startsWith("line") ? false : undefined;
  } else if (kind === "radar") {
    const style = attr(kid(typeEl, "radarStyle"), "val");
    radarStyle = style === "filled" ? "filled" : style === "marker" || anyMarker ? "marker" : "standard";
  }

  const title = richText(kid(chart, "title", "tx", "rich"));

  // legend position + chart/plot area styling
  const legendEl = kid(chart, "legend");
  const lp = attr(kid(legendEl, "legendPos"), "val");
  const legendPos = lp === "b" || lp === "t" || lp === "l" ? lp : undefined;
  const chartSpPr = kid(doc.documentElement, "spPr");
  const chartFill = chartSpPr ? parseFill(chartSpPr) : undefined;
  const chartBorder = chartSpPr ? parseLn(chartSpPr) : undefined;
  const plotSpPr = kid(plotArea, "spPr");
  const plotFill = plotSpPr ? parseFill(plotSpPr) : undefined;
  const plotBorder = plotSpPr ? parseLn(plotSpPr) : undefined;

  // chart-wide text size (c:txPr/a:p/a:pPr/a:defRPr@sz)
  let labelSizePt: number | undefined;
  const txPr = kid(doc.documentElement, "txPr");
  if (txPr) {
    const walkSz = (el: Element): void => {
      for (let i = 0; i < el.childNodes.length; i++) {
        const c = el.childNodes[i] as Element;
        if (c.nodeType !== 1) continue;
        if (c.localName === "defRPr") {
          const sz = iattr(c, "sz", 0);
          if (sz) labelSizePt = sz / 100;
          return;
        }
        walkSz(c);
      }
    };
    walkSz(txPr);
  }

  return {
    chart: kind, categories, series, title,
    legend: !!legendEl, legendPos,
    grouping, marker, smooth, radarStyle, labelSizePt,
    dataLabels, errorBarsPct, axisTitleX, axisTitleY, hideAxisX, hideAxisY,
    chartFill, chartBorder, plotFill, plotBorder, markerSizePt,
    pointColors, gridColor, hideGridlines,
  };
}

// ---------- placeholders ----------
interface PhInfo {
  type: string | null;
  idx: string | null;
  x?: number; y?: number; w?: number; h?: number;
  /** Placeholder's own text style chain (txBody lstStyle) for size/font inheritance. */
  lst?: LstDefaults;
}

function phOf(shapeEl: Element): { type: string | null; idx: string | null } | null {
  const ph = kid(shapeEl, "nvSpPr", "nvPr", "ph");
  if (!ph) return null;
  return { type: attr(ph, "type"), idx: attr(ph, "idx") };
}

function collectPhs(spTree: Element | null): PhInfo[] {
  const out: PhInfo[] = [];
  if (!spTree) return out;
  for (const sp of kids(spTree, "sp")) {
    const ph = phOf(sp);
    if (!ph) continue;
    const info: PhInfo = {
      type: ph.type, idx: ph.idx,
      lst: parseLstStyle(kid(sp, "txBody", "lstStyle")),
    };
    const xfrm = kid(sp, "spPr", "xfrm");
    if (xfrm) {
      const off = kid(xfrm, "off"), ext = kid(xfrm, "ext");
      info.x = iattr(off, "x"); info.y = iattr(off, "y");
      info.w = iattr(ext, "cx"); info.h = iattr(ext, "cy");
    }
    out.push(info);
  }
  return out;
}

function findPh(phs: PhInfo[], type: string | null, idx: string | null): PhInfo | undefined {
  const normType = (t: string | null) => (t === "ctrTitle" ? "title" : t ?? "body");
  return (
    phs.find(p => p.idx !== null && p.idx === idx && (type === null || normType(p.type) === normType(type))) ??
    phs.find(p => p.idx !== null && idx !== null && p.idx === idx) ??
    phs.find(p => normType(p.type) === normType(type))
  );
}

interface ThemeFonts { major: string; minor: string }

// ---------- main ----------
export interface ParsedPptx {
  pres: Presentation;
  media: Map<string, MediaItem>;
  warnings: string[];
}

/**
 * Resource limits for parsing UNTRUSTED .pptx input (e.g. user uploads fetched
 * from object storage). Bounds memory/CPU so a malicious or malformed file can
 * only ever fail loudly, never hang or OOM the tab. Tune via the export.
 */
export const PARSE_LIMITS = {
  maxCompressedBytes: 150 * 1024 * 1024, // reject the raw .pptx above this
  maxTotalUncompressedBytes: 600 * 1024 * 1024, // sum of all extracted parts
  maxEntryUncompressedBytes: 200 * 1024 * 1024, // any single part
  maxEntries: 5000,            // zip member count
  maxSlides: 1000,             // slides actually parsed
  maxGroupDepth: 64,           // nested grpSp recursion
};

/** Reject zip bombs up front using the central-directory declared sizes (no decompression). */
function preflightZip(zip: JSZip): void {
  const names = Object.keys(zip.files);
  if (names.length > PARSE_LIMITS.maxEntries) {
    throw new Error(`This file has too many parts (${names.length}); refusing to open it.`);
  }
  let declared = 0;
  for (const name of names) {
    // _data.uncompressedSize is JSZip's parsed central-directory size
    const usize = (zip.files[name] as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;
    if (usize > PARSE_LIMITS.maxEntryUncompressedBytes) {
      throw new Error(`A part of this file is too large to open safely (${Math.round(usize / 1048576)} MB).`);
    }
    declared += usize;
  }
  if (declared > PARSE_LIMITS.maxTotalUncompressedBytes) {
    throw new Error(`This file expands too large to open safely (${Math.round(declared / 1048576)} MB).`);
  }
}

export async function parsePptx(data: ArrayBuffer | Uint8Array, fileName = "Presentation"): Promise<ParsedPptx> {
  const warnings: string[] = [];
  const compressedLen = data instanceof Uint8Array ? data.byteLength : data.byteLength;
  if (compressedLen > PARSE_LIMITS.maxCompressedBytes) {
    throw new Error(`This file is too large to open (${Math.round(compressedLen / 1048576)} MB).`);
  }
  const zip = await JSZip.loadAsync(data);
  preflightZip(zip);

  // Defence-in-depth: declared sizes can lie, so also meter ACTUAL extracted
  // bytes and abort if a part decompresses past the budget (true zip-bomb stop).
  let extractedBytes = 0;
  const meter = (n: number) => {
    extractedBytes += n;
    if (extractedBytes > PARSE_LIMITS.maxTotalUncompressedBytes) {
      throw new Error("This file expands too large to open safely; aborting.");
    }
  };

  // package rels -> presentation part
  const pkgRels = await parseRels(zip, "");
  let presPath = "ppt/presentation.xml";
  for (const r of pkgRels.values()) {
    if (r.type.endsWith("/officeDocument")) { presPath = r.target; break; }
  }
  const presFile = zip.file(presPath);
  if (!presFile) throw new Error("Not a PowerPoint file: missing " + presPath);
  const presDoc = parseXml(await presFile.async("text"));
  const presRoot = presDoc.documentElement;
  const presRels = await parseRels(zip, presPath);

  const sldSz = kid(presRoot, "sldSz");
  const slideWidth = iattr(sldSz, "cx", 12192000);
  const slideHeight = iattr(sldSz, "cy", 6858000);

  // theme (via presentation rels, else first theme part)
  let theme: ColorTheme = { ...OFFICE_THEME };
  let themePath: string | undefined;
  for (const r of presRels.values()) if (r.type.endsWith("/theme")) themePath = r.target;
  if (!themePath) {
    zip.forEach(p => { if (/ppt\/theme\/theme\d+\.xml$/.test(p) && !themePath) themePath = p; });
  }
  if (themePath && zip.file(themePath)) {
    try {
      theme = parseTheme(parseXml(await zip.file(themePath)!.async("text")), theme);
    } catch { warnings.push("Could not parse theme; using Office defaults."); }
  }
  const themeFonts: ThemeFonts = { major: theme.majorFont, minor: theme.minorFont };

  // layout/master placeholder caches
  interface PartStyles {
    phs: PhInfo[];
    /** slideMaster p:txStyles — base text styles for title/body/other placeholders. */
    txStyles?: { title?: LstDefaults; body?: LstDefaults; other?: LstDefaults };
  }
  const phCache = new Map<string, PartStyles>();
  async function phsForPart(path: string): Promise<PartStyles> {
    if (phCache.has(path)) return phCache.get(path)!;
    const f = zip.file(path);
    if (!f) { const empty = { phs: [] }; phCache.set(path, empty); return empty; }
    const doc = parseXml(await f.async("text"));
    const tree = kid(doc.documentElement, "cSld", "spTree");
    const out: PartStyles = { phs: collectPhs(tree) };
    const tx = kid(doc.documentElement, "txStyles");
    if (tx) {
      out.txStyles = {
        title: parseLstStyle(kid(tx, "titleStyle")),
        body: parseLstStyle(kid(tx, "bodyStyle")),
        other: parseLstStyle(kid(tx, "otherStyle")),
      };
    }
    phCache.set(path, out);
    return out;
  }

  // media registry
  const media = new Map<string, MediaItem>();
  const mediaByPath = new Map<string, string>(); // zip path -> media id
  async function mediaFor(path: string): Promise<string | null> {
    if (mediaByPath.has(path)) return mediaByPath.get(path)!;
    const f = zip.file(path);
    if (!f) return null;
    const bytes = await f.async("uint8array");
    meter(bytes.length);
    const mime = extToMime(path);
    const id = nextId("media");
    media.set(id, { id, mime, bytes, dataUrl: bytesToDataUrl(bytes, mime) });
    mediaByPath.set(path, id);
    return id;
  }

  /** Vector media (asvg:svgBlip): keep the svg for display/export plus its bitmap fallback. */
  async function mediaForSvg(svgPath: string, fallbackPath: string | null): Promise<string | null> {
    if (mediaByPath.has(svgPath)) return mediaByPath.get(svgPath)!;
    const sf = zip.file(svgPath);
    if (!sf) return fallbackPath ? mediaFor(fallbackPath) : null;
    let svgBytes = await sf.async("uint8array");
    meter(svgBytes.length);
    try {
      // ensure root width/height so the editor's <image> can render it
      const doc = parseXml(new TextDecoder().decode(svgBytes));
      const root = doc.documentElement;
      if (root.localName === "svg" && (!root.getAttribute("width") || !root.getAttribute("height"))) {
        const vb = (root.getAttribute("viewBox") ?? "").trim().split(/[\s,]+/).map(Number);
        const w = vb.length === 4 && vb[2] > 0 ? vb[2] : 300;
        const h = vb.length === 4 && vb[3] > 0 ? vb[3] : 300;
        root.setAttribute("width", String(w));
        root.setAttribute("height", String(h));
        if (XMLSerializerImpl) {
          svgBytes = new TextEncoder().encode(new XMLSerializerImpl().serializeToString(root as unknown as Node));
        }
      }
    } catch { /* keep bytes as-is */ }
    let pngFallback: Uint8Array | undefined;
    if (fallbackPath) {
      const pf = zip.file(fallbackPath);
      if (pf) { pngFallback = await pf.async("uint8array"); meter(pngFallback.length); }
    }
    const id = nextId("media");
    media.set(id, {
      id, mime: "image/svg+xml", bytes: svgBytes,
      dataUrl: bytesToDataUrl(svgBytes, "image/svg+xml"), pngFallback,
    });
    mediaByPath.set(svgPath, id);
    return id;
  }

  // slides in sldIdLst order
  const slideEntries: string[] = [];
  const sldIdLst = kid(presRoot, "sldIdLst");
  for (const sldId of kids(sldIdLst, "sldId")) {
    const rid = sldId.getAttributeNS?.("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
      ?? attr(sldId, "r:id");
    const rel = rid ? presRels.get(rid) : undefined;
    if (rel) slideEntries.push(rel.target);
  }
  if (!slideEntries.length) {
    zip.forEach(p => { if (/^ppt\/slides\/slide\d+\.xml$/.test(p)) slideEntries.push(p); });
    slideEntries.sort((a, b) => parseInt(a.match(/\d+/)?.[0] ?? "0") - parseInt(b.match(/\d+/)?.[0] ?? "0"));
  }

  if (slideEntries.length > PARSE_LIMITS.maxSlides) {
    warnings.push(`Only the first ${PARSE_LIMITS.maxSlides} of ${slideEntries.length} slides were opened.`);
    slideEntries.length = PARSE_LIMITS.maxSlides;
  }

  const slides: SlideModel[] = [];
  for (const slidePath of slideEntries) {
    const f = zip.file(slidePath);
    if (!f) continue;
    const doc = parseXml(await f.async("text"));
    const root = doc.documentElement;
    const slideRels = await parseRels(zip, slidePath);

    // layout + master placeholder chains
    let layoutPath: string | undefined;
    for (const r of slideRels.values()) if (r.type.endsWith("/slideLayout")) layoutPath = r.target;
    let layoutPhs: PhInfo[] = [], masterPhs: PhInfo[] = [];
    let masterTxStyles: PartStyles["txStyles"];
    if (layoutPath) {
      const layoutPart = await phsForPart(layoutPath);
      layoutPhs = layoutPart.phs;
      const layoutRels = await parseRels(zip, layoutPath);
      for (const r of layoutRels.values()) {
        if (r.type.endsWith("/slideMaster")) {
          const masterPart = await phsForPart(r.target);
          masterPhs = masterPart.phs;
          masterTxStyles = masterPart.txStyles;
        }
      }
    }

    const slide: SlideModel = { id: nextId("slide"), shapes: [] };

    // background (image backgrounds load their media like any picture)
    const bgPr = kid(root, "cSld", "bg", "bgPr");
    if (bgPr) {
      const f2 = await resolveImageFill(parseFill(bgPr));
      if (f2) slide.background = f2;
    }

    // transition (subset: fade / push family; others approximate to fade)
    const trans = kid(root, "transition");
    if (trans) {
      const spdRaw = attr(trans, "spd");
      const speed = spdRaw === "slow" ? "slow" : spdRaw === "fast" ? "fast" : "med";
      for (let i = 0; i < trans.childNodes.length; i++) {
        const n = trans.childNodes[i] as Element;
        if (n.nodeType !== 1) continue;
        const ln = n.localName;
        if (ln === "fade" || ln === "dissolve" || ln === "circle" || ln === "diamond") {
          slide.transition = { type: "fade", speed };
        } else if (ln === "push" || ln === "wipe" || ln === "cover" || ln === "pull") {
          const d = attr(n, "dir");
          slide.transition = { type: "push", dir: d === "r" ? "r" : d === "u" ? "u" : d === "d" ? "d" : "l", speed };
        } else if (ln === "cut") {
          slide.transition = { type: "none", speed };
        } else {
          slide.transition = { type: "fade", speed };
        }
        break;
      }
    }

    const spTree = kid(root, "cSld", "spTree");
    if (spTree) {
      await walkTree(spTree, { sx: 1, sy: 1, tx: 0, ty: 0 }, 0);
    }

    /** Swap a blipFill's r:embed placeholder for a loaded media id. */
    async function resolveImageFill(f: Fill | undefined): Promise<Fill | undefined> {
      if (f?.kind !== "image") return f;
      const rel = slideRels.get(f.mediaId);
      const id = rel ? await mediaFor(rel.target) : null;
      return id ? { ...f, mediaId: id } : { kind: "solid", color: { kind: "srgb", hex: "D9D9D9" } };
    }

    async function walkTree(tree: Element, t: { sx: number; sy: number; tx: number; ty: number; groupFill?: Fill; gid?: string }, depth: number) {
      // bound nested-group recursion so a pathologically deep file can't blow the stack
      if (depth > PARSE_LIMITS.maxGroupDepth) {
        if (!warnings.some(w => w.startsWith("Deeply nested groups"))) {
          warnings.push("Deeply nested groups were truncated.");
        }
        return;
      }
      for (let i = 0; i < tree.childNodes.length; i++) {
        const n = tree.childNodes[i] as Element;
        if (n.nodeType !== 1) continue;
        switch (n.localName) {
          case "sp": await handleSp(n, t); break;
          case "cxnSp": await handleSp(n, t, true); break;
          case "pic": await handlePic(n, t); break;
          case "AlternateContent": {
            // PowerPoint wraps modern pics (SVG, 3D, effects) in mc:AlternateContent —
            // skipping it silently drops images. Prefer the Fallback branch (always
            // plain DrawingML); use Choice when no Fallback exists (we read svgBlip).
            const branch = kid(n, "Fallback") ?? kid(n, "Choice");
            if (branch) await walkTree(branch, t, depth);
            break;
          }
          case "grpSp": {
            const grpSpPr = kid(n, "grpSpPr");
            const xfrm = kid(grpSpPr, "xfrm");
            const off = kid(xfrm, "off"), ext = kid(xfrm, "ext");
            const chOff = kid(xfrm, "chOff"), chExt = kid(xfrm, "chExt");
            const extCx = iattr(ext, "cx", 1) || 1, extCy = iattr(ext, "cy", 1) || 1;
            const chCx = iattr(chExt, "cx", extCx) || extCx, chCy = iattr(chExt, "cy", extCy) || extCy;
            const kx = extCx / chCx, ky = extCy / chCy;
            // geometry flattens, but membership is kept: every shape that comes out of
            // this (outermost) group gets a shared groupId so it stays grouped here
            const isOutermost = !t.gid;
            const gid = t.gid ?? nextId("grp");
            const startLen = slide.shapes.length;
            await walkTree(n, {
              sx: t.sx * kx,
              sy: t.sy * ky,
              tx: t.tx + (iattr(off, "x") - iattr(chOff, "x") * kx) * t.sx,
              ty: t.ty + (iattr(off, "y") - iattr(chOff, "y") * ky) * t.sy,
              groupFill: parseFill(grpSpPr, t.groupFill) ?? t.groupFill,
              gid,
            }, depth + 1);
            if (isOutermost) {
              for (let k = startLen; k < slide.shapes.length; k++) {
                slide.shapes[k] = { ...slide.shapes[k], groupId: gid };
              }
            }
            break;
          }
          case "graphicFrame": {
            const xfrm = kid(n, "xfrm");
            const off = kid(xfrm, "off"), ext = kid(xfrm, "ext");
            const name = attr(kid(n, "nvGraphicFramePr", "cNvPr"), "name") ?? "Object";
            const fx = t.tx + iattr(off, "x") * t.sx, fy = t.ty + iattr(off, "y") * t.sy;
            const fw = iattr(ext, "cx") * t.sx, fh = iattr(ext, "cy") * t.sy;
            const gd = kid(n, "graphic", "graphicData");

            const tbl = kid(gd, "tbl");
            if (tbl) {
              const dflt: TextDefaults = { sizePt: 14, font: "+mn-lt", color: { kind: "scheme", slot: "dk1" } };
              slide.shapes.push({
                kind: "table", id: nextId("tbl"), name,
                x: Math.round(fx), y: Math.round(fy), w: Math.round(fw), h: Math.round(fh), rot: 0,
                ...parseTableEl(tbl, dflt, themeFonts),
              });
              break;
            }

            const chartRefEl = kid(gd, "chart");
            if (chartRefEl) {
              const rid = chartRefEl.getAttributeNS?.("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
                ?? attr(chartRefEl, "r:id");
              const rel = rid ? slideRels.get(rid) : undefined;
              const file = rel ? zip.file(rel.target) : null;
              if (file) {
                try {
                  const parsed = parseChartDoc(parseXml(await file.async("text")));
                  if (parsed) {
                    slide.shapes.push({
                      kind: "chart", id: nextId("chart"), name,
                      x: Math.round(fx), y: Math.round(fy), w: Math.round(fw), h: Math.round(fh), rot: 0,
                      ...parsed,
                    });
                    break;
                  }
                  warnings.push(`Slide ${slides.length + 1}: chart type in '${name}' not supported — imported as placeholder.`);
                } catch {
                  warnings.push(`Slide ${slides.length + 1}: could not parse chart '${name}'.`);
                }
              }
            }

            slide.shapes.push(placeholderRect(name + " (object not supported)", fx, fy, fw, fh));
            if (!chartRefEl) warnings.push(`Slide ${slides.length + 1}: '${name}' is an embedded object — imported as placeholder.`);
            break;
          }
        }
      }
    }

    async function handleSp(n: Element, t: { sx: number; sy: number; tx: number; ty: number; groupFill?: Fill }, isCxn = false) {
      const nv = isCxn ? kid(n, "nvCxnSpPr", "cNvPr") : kid(n, "nvSpPr", "cNvPr");
      const name = attr(nv, "name") ?? "Shape";
      const spPr = kid(n, "spPr");
      const xfrm = kid(spPr, "xfrm");
      const ph = isCxn ? null : phOf(n);

      const layoutPh = ph ? findPh(layoutPhs, ph.type, ph.idx) : undefined;
      const masterPh = ph ? findPh(masterPhs, ph.type, ph.idx) : undefined;

      let x = 0, y = 0, w = 0, h = 0, rot = 0, flipH = false, flipV = false;
      if (xfrm) {
        const off = kid(xfrm, "off"), ext = kid(xfrm, "ext");
        x = iattr(off, "x"); y = iattr(off, "y");
        w = iattr(ext, "cx"); h = iattr(ext, "cy");
        rot = iattr(xfrm, "rot", 0) / 60000;
        flipH = attr(xfrm, "flipH") === "1";
        flipV = attr(xfrm, "flipV") === "1";
      } else if (ph) {
        const inh = (layoutPh?.w !== undefined ? layoutPh : undefined)
          ?? (masterPh?.w !== undefined ? masterPh : undefined);
        if (inh && inh.x !== undefined) { x = inh.x; y = inh.y!; w = inh.w!; h = inh.h!; }
        else { x = slideWidth * 0.1; y = slideHeight * 0.1; w = slideWidth * 0.8; h = slideHeight * 0.2; }
      } else {
        return; // no geometry at all
      }
      x = t.tx + x * t.sx; y = t.ty + y * t.sy; w *= t.sx; h *= t.sy;

      const prstGeomEl = kid(spPr, "prstGeom");
      const prst = attr(prstGeomEl, "prst");
      const custGeom = kid(spPr, "custGeom");
      let geom: PresetGeom = "rect";
      let adj: Record<string, number> | undefined;
      let custPath: SpShape["custPath"];
      let rawGeomXml: string | undefined;
      if (prst) {
        if (hasPreset(prst)) {
          geom = prst;
          adj = parseAdj(prstGeomEl);
        } else {
          warnings.push(`Unsupported shape preset '${prst}' rendered as rectangle.`);
        }
      } else if (custGeom) {
        custPath = parseCustGeom(custGeom) ?? undefined;
        if (custPath && XMLSerializerImpl) {
          try { rawGeomXml = new XMLSerializerImpl().serializeToString(custGeom as unknown as Node); } catch { /* render-only */ }
        }
        if (!custPath) warnings.push(`Custom geometry on '${name}' rendered as rectangle.`);
      } else if (isCxn) {
        geom = "line";
      }

      const isTitle = ph?.type === "title" || ph?.type === "ctrTitle";
      const isBodyPh = !!ph && !isTitle && ph.type !== "dt" && ph.type !== "ftr" && ph.type !== "sldNum";
      const dflt: TextDefaults = {
        sizePt: isTitle ? 44 : isBodyPh ? 18 : 18,
        font: isTitle ? "+mj-lt" : "+mn-lt",
        color: { kind: "scheme", slot: "dk1" },
        align: isTitle && ph?.type === "ctrTitle" ? "ctr" : undefined,
        bullet: false,
      };
      // placeholder text-style chain: master txStyles <- master ph <- layout ph
      // (the shape's own lstStyle merges on top inside parseTxBody)
      const masterBase = ph && masterTxStyles
        ? (isTitle ? masterTxStyles.title : isBodyPh ? masterTxStyles.body : masterTxStyles.other)
        : undefined;
      const inheritedLst = ph ? mergeLst(masterBase, masterPh?.lst, layoutPh?.lst) : undefined;
      const text = parseTxBody(kid(n, "txBody"), dflt, themeFonts, inheritedLst);

      const explicitFill = parseFill(spPr, t.groupFill);
      const fill: Fill = explicitFill ?? (ph || attr(kid(n, isCxn ? "nvCxnSpPr" : "nvSpPr", "cNvSpPr"), "txBox") === "1"
        ? { kind: "none" }
        : kid(n, "style", "fillRef")
          ? { kind: "solid", color: parseColorChoice(kid(n, "style", "fillRef")) ?? { kind: "scheme", slot: "accent1" } }
          : { kind: "none" });
      const line: LineProps = parseLn(spPr) ?? (kid(n, "style", "lnRef")
        ? { fill: { kind: "solid", color: parseColorChoice(kid(n, "style", "lnRef")) ?? { kind: "scheme", slot: "accent1", lumMod: 75 } }, widthPt: 1 }
        : { fill: { kind: "none" }, widthPt: 1 });

      const sp: SpShape = {
        kind: "sp", id: nextId("sp"), name,
        x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h),
        rot, flipH: flipH || undefined, flipV: flipV || undefined,
        geom, adj, custPath, rawGeomXml,
        isTextBox: !!ph || attr(kid(n, "nvSpPr", "cNvSpPr"), "txBox") === "1" || undefined,
        fill: (await resolveImageFill(fill))!, line, text,
      };
      slide.shapes.push(sp);
    }

    async function handlePic(n: Element, t: { sx: number; sy: number; tx: number; ty: number }) {
      const name = attr(kid(n, "nvPicPr", "cNvPr"), "name") ?? "Picture";
      const xfrm = kid(n, "spPr", "xfrm");
      const off = kid(xfrm, "off"), ext = kid(xfrm, "ext");
      const blip = kid(n, "blipFill", "blip");
      const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
      const embed = blip?.getAttributeNS?.(NS_REL, "embed") ?? attr(blip, "r:embed");
      // PowerPoint vector images: a:blip/a:extLst/a:ext/asvg:svgBlip alongside the bitmap fallback
      let svgEmbed: string | null = null;
      for (const ext of kids(kid(blip, "extLst"), "ext")) {
        const sb = kid(ext, "svgBlip");
        if (sb) {
          svgEmbed = sb.getAttributeNS?.(NS_REL, "embed") ?? attr(sb, "r:embed");
          break;
        }
      }
      let mediaId: string | null = null;
      if (svgEmbed) {
        const srel = slideRels.get(svgEmbed);
        const frel = embed ? slideRels.get(embed) : undefined;
        if (srel) mediaId = await mediaForSvg(srel.target, frel?.target ?? null);
      }
      if (!mediaId && embed) {
        const rel = slideRels.get(embed);
        if (rel) mediaId = await mediaFor(rel.target);
      }
      if (!mediaId) { warnings.push(`Image '${name}' could not be loaded.`); return; }
      // crop (a:srcRect, 1000ths of a percent per edge)
      const srcRectEl = kid(n, "blipFill", "srcRect");
      const srcRect = srcRectEl ? {
        l: iattr(srcRectEl, "l", 0) / 100000,
        t: iattr(srcRectEl, "t", 0) / 100000,
        r: iattr(srcRectEl, "r", 0) / 100000,
        b: iattr(srcRectEl, "b", 0) / 100000,
      } : undefined;
      // frame geometry (rounded corners, circle crop, ...)
      const picGeomEl = kid(n, "spPr", "prstGeom");
      const picPrst = attr(picGeomEl, "prst");
      const picGeom = picPrst && picPrst !== "rect" && hasPreset(picPrst) ? picPrst : undefined;
      slide.shapes.push({
        kind: "pic", id: nextId("pic"), name, mediaId,
        x: Math.round(t.tx + iattr(off, "x") * t.sx),
        y: Math.round(t.ty + iattr(off, "y") * t.sy),
        w: Math.round(iattr(ext, "cx") * t.sx),
        h: Math.round(iattr(ext, "cy") * t.sy),
        rot: iattr(xfrm, "rot", 0) / 60000,
        flipH: attr(xfrm, "flipH") === "1" || undefined,
        flipV: attr(xfrm, "flipV") === "1" || undefined,
        geom: picGeom,
        adj: picGeom ? parseAdj(picGeomEl) : undefined,
        srcRect: srcRect && (srcRect.l || srcRect.t || srcRect.r || srcRect.b) ? srcRect : undefined,
      });
    }

    // speaker notes (notesSlide part -> body placeholder text)
    for (const r of slideRels.values()) {
      if (!r.type.endsWith("/notesSlide")) continue;
      const nf = zip.file(r.target);
      if (!nf) continue;
      try {
        const ndoc = parseXml(await nf.async("text"));
        const tree = kid(ndoc.documentElement, "cSld", "spTree");
        const lines: string[] = [];
        for (const sp of kids(tree, "sp")) {
          const ph = kid(sp, "nvSpPr", "nvPr", "ph");
          if (attr(ph, "type") !== "body") continue;
          for (const p of kids(kid(sp, "txBody"), "p")) {
            const parts: string[] = [];
            for (const rEl of kids(p, "r")) parts.push(kid(rEl, "t")?.textContent ?? "");
            lines.push(parts.join(""));
          }
        }
        const text = lines.join("\n").replace(/\n+$/, "");
        if (text.trim()) slide.notes = text;
      } catch { /* notes are best-effort */ }
    }

    slides.push(slide);
  }

  if (!slides.length) throw new Error("No slides found in file");

  const title = fileName.replace(/\.pptx?$/i, "");
  return { pres: { slideWidth, slideHeight, slides, theme, title }, media, warnings };
}

function placeholderRect(name: string, x: number, y: number, w: number, h: number): SpShape {
  return {
    kind: "sp", id: nextId("sp"), name, x, y, w, h, rot: 0, geom: "rect",
    fill: { kind: "solid", color: { kind: "srgb", hex: "F2F2F2" } },
    line: { fill: { kind: "solid", color: { kind: "srgb", hex: "BFBFBF" } }, widthPt: 1, dash: "dash" },
    text: {
      paragraphs: [{ runs: [{ text: name, sizePt: 12, font: "Arial", color: { kind: "srgb", hex: "808080" } }], align: "ctr", bullet: "none", level: 0 }],
      anchor: "ctr", wrap: true, insets: [91440, 45720, 91440, 45720],
    },
  };
}

function parseTheme(doc: Document, base: ColorTheme): ColorTheme {
  const els = kid(doc.documentElement, "themeElements");
  const scheme = kid(els, "clrScheme");
  const out = { ...base, name: attr(scheme, "name") ?? base.name };
  if (scheme) {
    const grab = (local: string): string | null => {
      const el = kids(scheme, local)[0];
      if (!el) return null;
      const srgb = kid(el, "srgbClr");
      if (srgb) return (attr(srgb, "val") ?? "").toUpperCase() || null;
      const sys = kid(el, "sysClr");
      if (sys) return (attr(sys, "lastClr") ?? "").toUpperCase() || null;
      return null;
    };
    for (const slot of ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"] as const) {
      const v = grab(slot);
      if (v) out[slot] = v;
    }
  }
  const fonts = kid(els, "fontScheme");
  const major = attr(kid(fonts, "majorFont", "latin"), "typeface");
  const minor = attr(kid(fonts, "minorFont", "latin"), "typeface");
  if (major) out.majorFont = major;
  if (minor) out.minorFont = minor;
  return out;
}
