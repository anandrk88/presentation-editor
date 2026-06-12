import type {
  ChartKind, ChartShape, ColorRef, ColorTheme, Fill, LineProps, Paragraph,
  PresetGeom, Presentation, Run, SchemeSlot, Shape, SlideModel, SpShape,
  TableCell, TableShape, TextBody,
} from "./types";

export const EMU_PER_INCH = 914400;
export const EMU_PER_PX = 9525; // 96 dpi
export const EMU_PER_PT = 12700;

// 16:9 default slide size (same as PowerPoint / OnlyOffice default)
export const SLIDE_W = 12192000;
export const SLIDE_H = 6858000;

let idCounter = 1;
export function nextId(prefix = "obj"): string {
  return `${prefix}_${Date.now().toString(36)}_${(idCounter++).toString(36)}`;
}

/** Standard "Office" theme (matches ECMA-376 default theme part). */
export const OFFICE_THEME: ColorTheme = {
  name: "Office",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "44546A",
  lt2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
  hlink: "0563C1",
  folHlink: "954F72",
  majorFont: "Calibri Light",
  minorFont: "Calibri",
};

/**
 * Dark-cover friendly scheme for imported report templates (light text slots,
 * blue accents) — patterns with scheme-color refs adapt to whichever is active.
 */
export const GROWTH_THEME: ColorTheme = {
  ...OFFICE_THEME,
  name: "Growth Report",
  dk1: "F4F7FF",
  dk2: "0B1240",
  lt2: "E8ECF6",
  accent1: "4F9CFF",
  accent2: "8FB8FF",
  accent3: "94A3D6",
  accent5: "6FA8FF",
};

/** Alternative schemes for the Design tab (modeled on stock Office themes). */
export const THEME_PRESETS: ColorTheme[] = [
  OFFICE_THEME,
  GROWTH_THEME,
  { ...OFFICE_THEME, name: "Ion", dk2: "1E5155", accent1: "B01513", accent2: "EA6312", accent3: "E6B729", accent4: "6AAC90", accent5: "5F9C9D", accent6: "9D5D9D", hlink: "8E58B6", folHlink: "7F6F6F" },
  { ...OFFICE_THEME, name: "Wisp", dk2: "335B74", accent1: "A53010", accent2: "DE7E18", accent3: "9F8351", accent4: "728653", accent5: "92AA4C", accent6: "6AA2B5", hlink: "FB4A18", folHlink: "5F5F5F" },
  { ...OFFICE_THEME, name: "Slice", dk2: "146194", accent1: "052F61", accent2: "A50E82", accent3: "14967C", accent4: "6A9E1F", accent5: "E87D37", accent6: "C62324", hlink: "0D2E46", folHlink: "356A95" },
  { ...OFFICE_THEME, name: "Berlin", dk2: "9D360E", accent1: "F09415", accent2: "C1B56B", accent3: "4BAAFA", accent4: "0E6FC9", accent5: "695E69", accent6: "69A761", hlink: "FFEE76", folHlink: "DB5353" },
  { ...OFFICE_THEME, name: "Damask", dk2: "2A5867", accent1: "9EB060", accent2: "D2C29D", accent3: "8C85A5", accent4: "7E48A5", accent5: "5E62A5", accent6: "3C8AB0", hlink: "F5DB81", folHlink: "B2BC83" },
];

/**
 * Theme font references, exactly as OOXML spells them: text that uses these
 * re-fonts automatically when the theme's heading/body fonts change.
 */
export const MAJOR_FONT = "+mj-lt"; // Headings
export const MINOR_FONT = "+mn-lt"; // Body

export function isThemeFont(font: string): boolean {
  return font === MAJOR_FONT || font === MINOR_FONT;
}

export function resolveFontName(font: string, theme: ColorTheme): string {
  if (font === MAJOR_FONT) return theme.majorFont;
  if (font === MINOR_FONT) return theme.minorFont;
  return font;
}

/** Design-tab font pairs (heading / body), like PowerPoint's theme fonts gallery. */
export const THEME_FONT_PAIRS: { name: string; major: string; minor: string }[] = [
  { name: "Office", major: "Calibri Light", minor: "Calibri" },
  { name: "Arial", major: "Arial", minor: "Arial" },
  { name: "Segoe UI", major: "Segoe UI", minor: "Segoe UI" },
  { name: "Georgia – Verdana", major: "Georgia", minor: "Verdana" },
  { name: "Cambria – Calibri", major: "Cambria", minor: "Calibri" },
  { name: "Trebuchet – Tahoma", major: "Trebuchet MS", minor: "Tahoma" },
  { name: "Times New Roman", major: "Times New Roman", minor: "Times New Roman" },
  { name: "Impact – Segoe UI", major: "Impact", minor: "Segoe UI" },
];

export const FONT_LIST = [
  "Arial", "Calibri", "Calibri Light", "Cambria", "Comic Sans MS", "Consolas",
  "Courier New", "Georgia", "Impact", "Segoe UI", "Tahoma", "Times New Roman",
  "Trebuchet MS", "Verdana",
];

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 96];

export const srgb = (hex: string): ColorRef => ({ kind: "srgb", hex });
export const scheme = (slot: SchemeSlot, lumMod?: number, lumOff?: number): ColorRef =>
  ({ kind: "scheme", slot, lumMod, lumOff });

export const solidFill = (color: ColorRef): Fill => ({ kind: "solid", color });
export const noFill = (): Fill => ({ kind: "none" });

export function defaultLine(): LineProps {
  return { fill: { kind: "solid", color: { kind: "scheme", slot: "accent1", lumMod: 75 } }, widthPt: 1 };
}

export function defaultRun(text = "", sizePt = 18): Run {
  return { text, sizePt, font: MINOR_FONT, color: { kind: "scheme", slot: "dk1" } };
}

export function defaultParagraph(runs: Run[] = [], align: Paragraph["align"] = "l"): Paragraph {
  return { runs, align, bullet: "none", level: 0 };
}

export function emptyTextBody(anchor: TextBody["anchor"] = "ctr"): TextBody {
  return {
    paragraphs: [defaultParagraph([], "ctr")],
    anchor,
    wrap: true,
    insets: [91440, 45720, 91440, 45720], // OOXML bodyPr defaults
  };
}

export function makeShape(geom: PresetGeom, x: number, y: number, w: number, h: number, name?: string): SpShape {
  const isLine = geom === "line" || geom === "straightConnector1";
  return {
    kind: "sp",
    id: nextId("sp"),
    name: name ?? geomDisplayName(geom),
    x, y, w, h,
    rot: 0,
    geom,
    fill: isLine ? noFill() : solidFill({ kind: "scheme", slot: "accent1" }),
    line: isLine
      ? { fill: { kind: "solid", color: { kind: "scheme", slot: "accent1" } }, widthPt: 2 }
      : defaultLine(),
    // empty-text style carrier: typed text inherits 18pt white (PowerPoint look on accent fill)
    text: isLine ? undefined : {
      ...emptyTextBody("ctr"),
      paragraphs: [defaultParagraph([{ ...defaultRun(""), color: { kind: "scheme", slot: "lt1" } }], "ctr")],
    },
  };
}

export function tableCell(text = "", style: Partial<Run> = {}, opts: Partial<TableCell> = {}): TableCell {
  // always carry one run, even when empty — it seeds the cell's text style
  const run: Run = {
    text,
    sizePt: style.sizePt ?? 14,
    font: style.font ?? MINOR_FONT,
    color: style.color ?? { kind: "scheme", slot: "dk1" },
    bold: style.bold,
    italic: style.italic,
  };
  return {
    text: {
      paragraphs: [{ runs: [run], align: "l", bullet: "none", level: 0 }],
      anchor: "ctr",
      wrap: true,
      insets: [91440, 45720, 91440, 45720],
    },
    ...opts,
  };
}

export function makeTable(rows: number, cols: number, x: number, y: number, w: number, h: number): TableShape {
  const cells: TableCell[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: TableCell[] = [];
    for (let c = 0; c < cols; c++) {
      row.push(r === 0
        ? tableCell("", { bold: true, color: { kind: "scheme", slot: "lt1" } })
        : tableCell(""));
    }
    cells.push(row);
  }
  return {
    kind: "table", id: nextId("tbl"), name: "Table",
    x, y, w, h, rot: 0,
    colW: Array(cols).fill(Math.round(w / cols)),
    rowH: Array(rows).fill(Math.round(h / rows)),
    cells,
    firstRow: true,
    bandRow: true,
  };
}

// ---------- table styles (PowerPoint Light/Medium/Dark × accent gallery) ----------
export interface TableCellStyle {
  fill: Fill;
  /** Text color the style would give this cell (applied when a style is chosen). */
  textColor: ColorRef;
  bold: boolean;
}

type TableStyleSource = Pick<TableShape, "firstRow" | "bandRow" | "totalRow" | "firstCol" | "lastCol" | "bandCol" | "styleFamily" | "accent" | "cells">;

/** Effective style for a table cell from the built-in style families. */
export function tableCellStyle(t: TableStyleSource, r: number, c: number): TableCellStyle {
  const family = t.styleFamily ?? "medium";
  const gray = t.accent === "none";
  const slot: SchemeSlot = t.accent && t.accent !== "none" ? t.accent : "accent1";
  const nRows = t.cells.length, nCols = t.cells[0]?.length ?? 0;
  const lt1: ColorRef = { kind: "scheme", slot: "lt1" };
  const dk1: ColorRef = { kind: "scheme", slot: "dk1" };
  const solid = (color: ColorRef): Fill => ({ kind: "solid", color });
  const accent = (lumMod?: number, lumOff?: number): ColorRef =>
    gray
      ? { kind: "scheme", slot: "dk1", lumMod: lumMod !== undefined ? Math.round(lumMod / 2) : undefined, lumOff: lumOff !== undefined ? Math.min(95, lumOff + 5) : undefined }
      : { kind: "scheme", slot, lumMod, lumOff };

  const isHeader = !!t.firstRow && r === 0;
  const isTotal = !!t.totalRow && r === nRows - 1 && !isHeader;
  const isFirstCol = !!t.firstCol && c === 0;
  const isLastCol = !!t.lastCol && c === nCols - 1;
  const bandStart = t.firstRow ? 1 : 0;
  const rowBanded = !!t.bandRow && !isHeader && !isTotal && (r - bandStart) % 2 === 0;
  const colBanded = !!t.bandCol && !isHeader && !isTotal && (c - (t.firstCol ? 1 : 0)) % 2 === 0;
  const emphasisCol = isFirstCol || isLastCol;

  if (family === "dark") {
    if (isHeader) return { fill: solid(gray ? dk1 : accent()), textColor: lt1, bold: true };
    if (isTotal) return { fill: solid({ kind: "scheme", slot: "dk1", lumOff: 35 }), textColor: lt1, bold: true };
    return {
      fill: solid({ kind: "scheme", slot: "dk1", lumOff: rowBanded || colBanded ? 28 : 16 }),
      textColor: lt1, bold: emphasisCol,
    };
  }

  if (family === "light") {
    if (isHeader || isTotal) return { fill: solid(lt1), textColor: gray ? dk1 : accent(), bold: true };
    if (rowBanded || colBanded) return { fill: solid(accent(20, 80)), textColor: dk1, bold: emphasisCol };
    return { fill: solid(lt1), textColor: dk1, bold: emphasisCol };
  }

  // medium — the classic "Medium Style 2" look
  if (isHeader) return { fill: solid(gray ? dk1 : accent()), textColor: lt1, bold: true };
  if (isTotal) return { fill: solid(accent(40, 60)), textColor: dk1, bold: true };
  if (rowBanded || colBanded) return { fill: solid(accent(20, 80)), textColor: dk1, bold: emphasisCol };
  return { fill: solid(lt1), textColor: dk1, bold: emphasisCol };
}

/** Default border color per family when the table carries no explicit override. */
export function tableBorderColor(t: Pick<TableShape, "styleFamily" | "accent" | "borderColor">): ColorRef {
  if (t.borderColor) return t.borderColor;
  const family = t.styleFamily ?? "medium";
  if (family === "light") {
    return t.accent && t.accent !== "none"
      ? { kind: "scheme", slot: t.accent, lumMod: 60, lumOff: 40 }
      : { kind: "scheme", slot: "dk1", lumMod: 35, lumOff: 65 };
  }
  if (family === "dark") return { kind: "scheme", slot: "lt1", alpha: 30 };
  return { kind: "scheme", slot: "lt1" };
}

/** Back-compat: old single-accent medium style fill (header / banded / plain). */
export function tableStyleFill(shape: Pick<TableShape, "firstRow" | "bandRow">, r: number): Fill {
  return tableCellStyle({ ...shape, cells: [[]] }, r, -1).fill;
}

export const CHART_NAMES: Record<ChartKind, string> = {
  column: "Column", bar: "Bar", line: "Line", pie: "Pie", doughnut: "Doughnut",
  area: "Area", scatter: "XY (Scatter)", radar: "Radar",
};

export function makeChart(kind: ChartKind, x: number, y: number, w: number, h: number, opts: Partial<ChartShape> = {}): ChartShape {
  const scatter = kind === "scatter";
  return {
    kind: "chart", id: nextId("chart"), name: `${CHART_NAMES[kind]} Chart`,
    x, y, w, h, rot: 0,
    chart: kind,
    categories: scatter ? ["0.7", "1.8", "2.6", "3.2", "4.1"] : ["Q1", "Q2", "Q3", "Q4"],
    series: scatter
      ? [{ name: "Series 1", values: [2.7, 3.2, 0.8, 1.2, 2.9] }]
      : [
          { name: "Series 1", values: [4.3, 2.5, 3.5, 4.5] },
          { name: "Series 2", values: [2.4, 4.4, 1.8, 2.8] },
        ],
    title: undefined,
    legend: true,
    marker: scatter ? true : undefined,
    ...opts,
  };
}

export function makeTextBox(x: number, y: number, w: number, h: number, text = ""): SpShape {
  return {
    kind: "sp",
    id: nextId("tx"),
    name: "TextBox",
    x, y, w, h,
    rot: 0,
    geom: "rect",
    isTextBox: true,
    fill: noFill(),
    line: { fill: noFill(), widthPt: 1 },
    text: {
      // always carry one run (possibly empty) so typed text inherits 18pt Arial,
      // not the browser UI font — see TextEditOverlay seed handling
      paragraphs: [defaultParagraph([defaultRun(text)], "l")],
      anchor: "t",
      wrap: true,
      insets: [91440, 45720, 91440, 45720],
    },
  };
}

function geomDisplayName(g: PresetGeom): string {
  const names: Partial<Record<PresetGeom, string>> = {
    rect: "Rectangle", roundRect: "Rounded Rectangle", ellipse: "Ellipse",
    triangle: "Triangle", rtTriangle: "Right Triangle", diamond: "Diamond",
    parallelogram: "Parallelogram", trapezoid: "Trapezoid", hexagon: "Hexagon",
    pentagon: "Pentagon", chevron: "Chevron", rightArrow: "Right Arrow",
    leftArrow: "Left Arrow", upArrow: "Up Arrow", downArrow: "Down Arrow",
    star5: "5-Point Star", heart: "Heart", line: "Line", straightConnector1: "Line",
  };
  return names[g] ?? "Shape";
}

/** Layout presets used when adding slides (mirrors PowerPoint's built-in layouts). */
export type LayoutKind = "title" | "titleContent" | "sectionHeader" | "twoContent" | "titleOnly" | "blank";

export const LAYOUT_NAMES: Record<LayoutKind, string> = {
  title: "Title Slide",
  titleContent: "Title and Content",
  sectionHeader: "Section Header",
  twoContent: "Two Content",
  titleOnly: "Title Only",
  blank: "Blank",
};

function placeholderBox(x: number, y: number, w: number, h: number, opts: {
  text?: string; sizePt: number; bold?: boolean; align?: Paragraph["align"];
  anchor?: TextBody["anchor"]; name: string; font?: string; bullets?: boolean;
}): SpShape {
  const para: Paragraph = {
    // run kept even when empty: it carries the placeholder's size/font so typing
    // into "Click to add title" yields 44pt, not the inherited UI font
    runs: [{ text: opts.text ?? "", sizePt: opts.sizePt, bold: opts.bold, font: opts.font ?? MAJOR_FONT, color: { kind: "scheme", slot: "dk1" } }],
    align: opts.align ?? "l",
    bullet: opts.bullets ? "char" : "none",
    level: 0,
  };
  return {
    kind: "sp",
    id: nextId("ph"),
    name: opts.name,
    x, y, w, h,
    rot: 0,
    geom: "rect",
    isTextBox: true,
    fill: noFill(),
    line: { fill: noFill(), widthPt: 1 },
    text: { paragraphs: [para], anchor: opts.anchor ?? "t", wrap: true, insets: [91440, 45720, 91440, 45720] },
  };
}

export function makeSlide(layout: LayoutKind = "titleContent"): SlideModel {
  const shapes: Shape[] = [];
  const W = SLIDE_W, H = SLIDE_H;
  const M = Math.round(W * 0.066); // side margin ~ PowerPoint defaults
  const CW = W - 2 * M;

  switch (layout) {
    case "title":
      shapes.push(placeholderBox(M, Math.round(H * 0.30), CW, Math.round(H * 0.21), {
        name: "Title 1", sizePt: 44, align: "ctr", anchor: "b", text: "",
      }));
      shapes.push(placeholderBox(Math.round(W * 0.17), Math.round(H * 0.54), Math.round(W * 0.66), Math.round(H * 0.12), {
        name: "Subtitle 2", sizePt: 20, align: "ctr", anchor: "t", font: MINOR_FONT,
      }));
      break;
    case "titleContent":
      shapes.push(placeholderBox(M, Math.round(H * 0.05), CW, Math.round(H * 0.15), {
        name: "Title 1", sizePt: 36, anchor: "ctr",
      }));
      shapes.push(placeholderBox(M, Math.round(H * 0.24), CW, Math.round(H * 0.66), {
        name: "Content Placeholder 2", sizePt: 20, font: MINOR_FONT, bullets: true,
      }));
      break;
    case "sectionHeader":
      shapes.push(placeholderBox(M, Math.round(H * 0.52), CW, Math.round(H * 0.16), {
        name: "Title 1", sizePt: 40, bold: true, anchor: "b",
      }));
      shapes.push(placeholderBox(M, Math.round(H * 0.70), CW, Math.round(H * 0.10), {
        name: "Text Placeholder 2", sizePt: 18, font: MINOR_FONT,
      }));
      break;
    case "twoContent": {
      shapes.push(placeholderBox(M, Math.round(H * 0.05), CW, Math.round(H * 0.15), {
        name: "Title 1", sizePt: 36, anchor: "ctr",
      }));
      const colW = Math.round((CW - Math.round(W * 0.02)) / 2);
      shapes.push(placeholderBox(M, Math.round(H * 0.24), colW, Math.round(H * 0.66), {
        name: "Content Placeholder 2", sizePt: 18, font: MINOR_FONT, bullets: true,
      }));
      shapes.push(placeholderBox(M + colW + Math.round(W * 0.02), Math.round(H * 0.24), colW, Math.round(H * 0.66), {
        name: "Content Placeholder 3", sizePt: 18, font: MINOR_FONT, bullets: true,
      }));
      break;
    }
    case "titleOnly":
      shapes.push(placeholderBox(M, Math.round(H * 0.05), CW, Math.round(H * 0.15), {
        name: "Title 1", sizePt: 36, anchor: "ctr",
      }));
      break;
    case "blank":
      break;
  }
  return { id: nextId("slide"), shapes };
}

export function newPresentation(): Presentation {
  return {
    slideWidth: SLIDE_W,
    slideHeight: SLIDE_H,
    slides: [makeSlide("title")],
    theme: OFFICE_THEME,
    title: "New Presentation",
  };
}

/** Resolve a ColorRef to a CSS color (hex, or rgba() when the ref carries alpha). */
export function resolveColor(c: ColorRef, theme: ColorTheme): string {
  let hex = c.kind === "srgb" ? c.hex : theme[c.slot] ?? "000000";
  if (c.lumMod !== undefined || c.lumOff !== undefined) {
    hex = applyLum(hex, c.lumMod ?? 100, c.lumOff ?? 0);
  }
  if (c.alpha !== undefined && c.alpha < 100) {
    const r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
    return `rgba(${r},${g},${b},${(c.alpha / 100).toFixed(3)})`;
  }
  return "#" + hex;
}

/** Approximate OOXML lumMod/lumOff (percent values) in HSL space. */
function applyLum(hex: string, lumMod: number, lumOff: number): string {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  let h = 0, s = 0;
  let l = (mx + mn) / 2;
  const d = mx - mn;
  if (d > 0) {
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    switch (mx) {
      case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
      case g: h = ((b - r) / d + 2) / 6; break;
      default: h = ((r - g) / d + 4) / 6;
    }
  }
  l = Math.min(1, Math.max(0, (l * lumMod) / 100 + lumOff / 100));
  const hue2rgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r2: number, g2: number, b2: number;
  if (s === 0) { r2 = g2 = b2 = l; }
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r2 = hue2rgb(p, q, h + 1 / 3);
    g2 = hue2rgb(p, q, h);
    b2 = hue2rgb(p, q, h - 1 / 3);
  }
  const to2 = (v: number) => Math.round(v * 255).toString(16).padStart(2, "0").toUpperCase();
  return to2(r2) + to2(g2) + to2(b2);
}

/** Flatten a fill to one CSS color (gradients -> middle-most stop). Renderer handles real gradients itself. */
export function resolveFill(f: Fill | undefined, theme: ColorTheme): string | "none" {
  if (!f || f.kind === "none") return "none";
  if (f.kind === "gradient") {
    if (!f.stops.length) return "none";
    const sorted = [...f.stops].sort((a, b) => a.pos - b.pos);
    return resolveColor(sorted[Math.floor(sorted.length / 2)].color, theme);
  }
  if (f.kind === "image") return "#D9D9D9";
  if (f.kind === "pattern") return resolveColor(f.fg, theme);
  return resolveColor(f.color, theme);
}
