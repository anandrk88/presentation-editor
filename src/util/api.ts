/**
 * Public scripting API — a stable, host-friendly surface over the editor.
 *
 * Lets a host application (the page embedding the editor, or your own feature
 * code) READ what's on screen — the active slide, the current selection, and
 * every element's properties (text, image, fill, geometry, table, chart) — and
 * MUTATE it programmatically (set text, swap an image, recolor, move/resize,
 * select, delete, undo/redo).
 *
 * Two ways to reach it:
 *   • Same window / same-origin iframe:  window.presentationEditor  (synchronous)
 *   • Cross-origin iframe:               postMessage { type: "pe:invoke", … }
 *     (see initEmbedBridge in embed.ts — every method below is callable by name)
 *
 * The shapes returned here are a DELIBERATELY STABLE projection of the internal
 * document model — they do not change when the model is refactored. Geometry is
 * reported in EMU (the native OOXML unit: 914400 per inch, 9525 per CSS px) with
 * a parallel `px` block for convenience.
 *
 * Element reads work on any slide; element WRITES target the active slide, which
 * mirrors the editor's own model. To edit another slide, selectSlide(i) first.
 */
import { store } from "../state/store";
import {
  EMU_PER_INCH, EMU_PER_PT, EMU_PER_PX, MINOR_FONT,
  makeChart, makeShape, makeTable, makeTextBox, nextId,
  resolveColor, resolveFontName,
} from "../model/defaults";
import { loadImageFromUrl } from "./loadImage";
import { exportAuthUrl, permissions, uiConfig, type Permissions, type UIConfig } from "./config";
import type { LayoutKind } from "../model/defaults";
import type {
  BulletKind, ChartKind, ChartShape, ColorRef, Fill, PicShape, PresetGeom, Run, SchemeSlot,
  Shape, SpShape, TableShape, TextAlign, TextBody,
} from "../model/types";

// ----------------------------- serialized shapes -----------------------------

export type ElementKind = "shape" | "image" | "table" | "chart";

export interface RunInfo {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  /** "all" = ALL CAPS, "small" = small caps (display effect). */
  caps?: "all" | "small";
  sizePt: number;
  /** Resolved font family (theme refs like +mn-lt already substituted). */
  font: string;
  /** Resolved CSS color string. */
  color: string;
}

export interface ParagraphInfo {
  /** The paragraph's text (all runs concatenated). */
  text: string;
  align: TextAlign;
  level: number;
  bullet: "none" | "char" | "num";
  runs: RunInfo[];
}

export type FillInfo =
  | { kind: "none" }
  | { kind: "solid"; color: string }
  | { kind: "gradient"; angle: number; stops: { pos: number; color: string }[] }
  | { kind: "image"; mediaId: string; tile: boolean }
  | { kind: "pattern"; preset: string; fg: string; bg: string };

export interface ImageInfo {
  mediaId: string;
  mime?: string;
  /** Data URL for the image bytes (renderable in an <img>); may be large. */
  dataUrl?: string;
  /** Crop fractions cut from each edge (0–1), if the picture is cropped. */
  crop?: { l: number; t: number; r: number; b: number };
}

export interface TableInfo {
  rows: number;
  cols: number;
  /** Plain text of every cell, [row][col]. */
  cells: string[][];
}

export interface ChartInfo {
  chartType: ChartShape["chart"];
  title?: string;
  legend: boolean;
  categories: string[];
  series: { name: string; values: number[] }[];
}

export interface ElementInfo {
  id: string;
  name: string;
  kind: ElementKind;
  /** Members of the same group share this id (move/resize together). */
  groupId?: string;
  // geometry — EMU (native) …
  x: number; y: number; w: number; h: number;
  rot: number; flipH: boolean; flipV: boolean;
  // … and the same rectangle in CSS px, for convenience.
  px: { x: number; y: number; w: number; h: number };
  /** Plain text (paragraphs joined by "\n") — text shapes only. */
  text?: string;
  /** Structured text — text shapes only. */
  paragraphs?: ParagraphInfo[];
  /** Preset geometry name (e.g. "rect", "ellipse", "hexagon") — shapes only. */
  geom?: string;
  /** Fill — shapes only (images/tables/charts have their own treatment). */
  fill?: FillInfo;
  /** Outline — shapes only. */
  line?: { color: string; widthPt: number; dash?: string };
  /** Picture details — images only. */
  image?: ImageInfo;
  /** Table details — tables only. */
  table?: TableInfo;
  /** Chart details — charts only. */
  chart?: ChartInfo;
}

export interface SlideInfo {
  index: number;
  id: string;
  /** First non-empty line of text on the slide, if any (a usable label). */
  title?: string;
  elementCount: number;
  /** Background as a fill (undefined = inherits the master/white). */
  background?: FillInfo;
  /** Present only when explicitly requested (getSlide) — the slide's elements. */
  elements?: ElementInfo[];
}

export interface SelectionInfo {
  slideIndex: number;
  /** Selected element ids on the active slide. */
  ids: string[];
  /** Full info for each selected element. */
  elements: ElementInfo[];
}

export interface DocumentInfo {
  title: string;
  slideCount: number;
  /** Slide size in EMU. */
  slideWidth: number;
  slideHeight: number;
  activeSlideIndex: number;
  dirty: boolean;
  canUndo: boolean;
  canRedo: boolean;
  /** Theme scheme colors as resolved hex (dk1, lt1, accent1…). */
  palette: Record<string, string>;
}

export type ApiEvent = "selection" | "slide" | "dirty" | "change";

// ----------------------------- serialization ---------------------------------

const toPx = (emu: number) => Math.round(emu / EMU_PER_PX);

function elementKind(s: Shape): ElementKind {
  return s.kind === "pic" ? "image" : s.kind === "table" ? "table" : s.kind === "chart" ? "chart" : "shape";
}

function describeFill(fill: Fill | undefined): FillInfo {
  const theme = store.pres.theme;
  if (!fill || fill.kind === "none") return { kind: "none" };
  if (fill.kind === "solid") return { kind: "solid", color: resolveColor(fill.color, theme) };
  if (fill.kind === "gradient")
    return { kind: "gradient", angle: fill.angle, stops: fill.stops.map(s => ({ pos: s.pos, color: resolveColor(s.color, theme) })) };
  if (fill.kind === "image") return { kind: "image", mediaId: fill.mediaId, tile: !!fill.tile };
  return { kind: "pattern", preset: fill.prst, fg: resolveColor(fill.fg, theme), bg: resolveColor(fill.bg, theme) };
}

/** Best-effort single CSS color for a fill (used for the outline color). */
function fillColorString(fill: Fill | undefined): string {
  const theme = store.pres.theme;
  if (!fill || fill.kind === "none") return "none";
  if (fill.kind === "solid") return resolveColor(fill.color, theme);
  if (fill.kind === "gradient") return fill.stops[0] ? resolveColor(fill.stops[0].color, theme) : "none";
  if (fill.kind === "pattern") return resolveColor(fill.fg, theme);
  return "none"; // image fill has no single color
}

function describeRun(r: Run): RunInfo {
  const theme = store.pres.theme;
  return {
    text: r.text,
    bold: !!r.bold, italic: !!r.italic, underline: !!r.underline, strike: !!r.strike,
    caps: r.caps,
    sizePt: r.sizePt,
    font: resolveFontName(r.font, theme),
    color: resolveColor(r.color, theme),
  };
}

function describeText(body: TextBody): { text: string; paragraphs: ParagraphInfo[] } {
  const paragraphs: ParagraphInfo[] = body.paragraphs.map(p => ({
    text: p.runs.map(r => r.text).join(""),
    align: p.align, level: p.level, bullet: p.bullet,
    runs: p.runs.map(describeRun),
  }));
  return { text: paragraphs.map(p => p.text).join("\n"), paragraphs };
}

/** Plain text of a table cell (its paragraphs joined). */
const cellText = (c: TableShape["cells"][number][number]) =>
  c.text.paragraphs.map(p => p.runs.map(r => r.text).join("")).join("\n");

export function describeElement(s: Shape): ElementInfo {
  const base: ElementInfo = {
    id: s.id, name: s.name, kind: elementKind(s), groupId: s.groupId,
    x: s.x, y: s.y, w: s.w, h: s.h, rot: s.rot, flipH: !!s.flipH, flipV: !!s.flipV,
    px: { x: toPx(s.x), y: toPx(s.y), w: toPx(s.w), h: toPx(s.h) },
  };
  if (s.kind === "sp") {
    const sp = s as SpShape;
    if (sp.text) { const t = describeText(sp.text); base.text = t.text; base.paragraphs = t.paragraphs; }
    base.geom = sp.geom;
    base.fill = describeFill(sp.fill);
    base.line = { color: fillColorString(sp.line.fill), widthPt: sp.line.widthPt, dash: sp.line.dash };
  } else if (s.kind === "pic") {
    const pic = s as PicShape;
    const m = store.media.get(pic.mediaId);
    base.image = { mediaId: pic.mediaId, mime: m?.mime, dataUrl: m?.dataUrl, crop: pic.srcRect };
  } else if (s.kind === "table") {
    const t = s as TableShape;
    base.table = { rows: t.rowH.length, cols: t.colW.length, cells: t.cells.map(row => row.map(cellText)) };
  } else if (s.kind === "chart") {
    const c = s as ChartShape;
    base.chart = { chartType: c.chart, title: c.title, legend: c.legend, categories: c.categories, series: c.series.map(se => ({ name: se.name, values: se.values })) };
  }
  return base;
}

function slideTitle(shapes: Shape[]): string | undefined {
  for (const s of shapes) {
    if (s.kind === "sp" && s.text) {
      const line = s.text.paragraphs.map(p => p.runs.map(r => r.text).join("")).join(" ").trim();
      if (line) return line.length > 80 ? line.slice(0, 80) + "…" : line;
    }
  }
  return undefined;
}

function describeSlide(index: number, withElements = false): SlideInfo {
  const sl = store.pres.slides[index];
  const info: SlideInfo = {
    index, id: sl.id, title: slideTitle(sl.shapes), elementCount: sl.shapes.length,
    background: sl.background ? describeFill(sl.background) : undefined,
  };
  if (withElements) info.elements = sl.shapes.map(describeElement);
  return info;
}

// ----------------------------- helpers ---------------------------------------

const activeShapes = (): Shape[] => store.currentSlide?.shapes ?? [];
const findOnActive = (id: string): Shape | undefined => activeShapes().find(s => s.id === id);

function normalizeHex(hex: string): string | null {
  const h = String(hex).replace(/^#/, "").trim().toUpperCase();
  if (/^[0-9A-F]{6}$/.test(h)) return h;
  if (/^[0-9A-F]{3}$/.test(h)) return h.split("").map(c => c + c).join("");
  return null;
}

/** A run to clone styling from when (re)building a text body. */
function templateRun(body: TextBody | undefined): Run {
  const r = body?.paragraphs.flatMap(p => p.runs)[0];
  return r ? { ...r } : { text: "", sizePt: 18, font: MINOR_FONT, color: { kind: "scheme", slot: "dk1" } };
}

const EMPTY_BODY = (): TextBody => ({
  paragraphs: [], anchor: "t", wrap: true, insets: [91440, 45720, 91440, 45720],
});

/** Box geometry for inserts (EMU). Omitted fields fall back to a centered default. */
export interface InsertBox { x?: number; y?: number; w?: number; h?: number }
export interface InsertTextOpts extends InsertBox { text?: string }
export interface InsertShapeOpts extends InsertBox { fillColor?: string }
export interface SeriesInput { name?: string; values: number[] }
/** Chart appearance options. Colors are hex strings ("4472C4" / "#4472C4");
 *  `null` clears a value (back to default/auto); invalid values are ignored. */
export interface ChartOptions {
  chartType?: ChartKind;                                   // change the chart's kind (column/bar/line/…)
  title?: string | null;                                   // null/"" clears
  legend?: boolean;                                        // show/hide legend
  legendPos?: "r" | "b" | "t" | "l";                       // right/bottom/top/left
  grouping?: "clustered" | "stacked" | "percentStacked";   // column/bar/area
  dataLabels?: boolean;                                    // value (or % for pie) labels
  dataLabelPos?: "outEnd" | "ctr" | "inEnd";               // outside / center / inside
  axisTitleX?: string | null;
  axisTitleY?: string | null;
  hideAxisX?: boolean;
  hideAxisY?: boolean;
  hideGridlines?: boolean;
  gridColor?: string | null;                               // hex
  markers?: boolean;                                       // line/scatter point markers
  markerSizePt?: number;
  smooth?: boolean;                                        // line/scatter curve smoothing
  radarStyle?: "standard" | "marker" | "filled";
  seriesColors?: (string | null)[];                        // per series; null = auto (accent cycle)
  sliceColors?: (string | null)[];                         // pie/doughnut per slice; null = auto
  chartFill?: string | null;                               // chart area: hex | "none" | null(clear)
  plotFill?: string | null;                                // plot area: hex | "none" | null(clear)
  labelSizePt?: number;                                    // global label font size
}
export interface InsertChartOpts extends InsertBox { categories?: string[]; series?: SeriesInput[]; options?: ChartOptions }
export interface ChartData { categories?: string[]; series?: SeriesInput[] }
export interface ParagraphStyle { align?: TextAlign; bullet?: BulletKind; level?: number }
export interface TextStyle { bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; sizePt?: number; font?: string; color?: string }

/** Resolve insert geometry: explicit EMU, else a centered default box. */
function placeBox(o: InsertBox, dw: number, dh: number): { x: number; y: number; w: number; h: number } {
  const w = Math.max(1, Math.round(o.w ?? dw)), h = Math.max(1, Math.round(o.h ?? dh));
  return {
    w, h,
    x: Math.round(o.x ?? (store.pres.slideWidth - w) / 2),
    y: Math.round(o.y ?? (store.pres.slideHeight - h) / 2),
  };
}

/** The 8 valid chart kinds — used to reject bad input before it can corrupt the .pptx. */
const CHART_KINDS: readonly ChartKind[] = ["column", "bar", "line", "area", "pie", "doughnut", "scatter", "radar"];

/** Apply ChartOptions onto a chart shape (returns a new shape). Invalid/unknown values are ignored. */
function applyChartOptions(ch: ChartShape, o: ChartOptions): ChartShape {
  const next: ChartShape = { ...ch };
  const ref = (v: string): ColorRef | undefined => { const h = normalizeHex(v); return h ? { kind: "srgb", hex: h } : undefined; };
  const fillOf = (v: string | null | undefined): Fill | undefined =>
    v == null ? undefined : v === "none" ? { kind: "none" } : (() => { const c = ref(v); return c ? { kind: "solid", color: c } : undefined; })();

  if (o.chartType && CHART_KINDS.includes(o.chartType)) next.chart = o.chartType;   // mirror the Chart-type dropdown
  if ("title" in o) next.title = o.title == null || o.title === "" ? undefined : String(o.title);
  if (typeof o.legend === "boolean") next.legend = o.legend;
  if (o.legendPos && ["r", "b", "t", "l"].includes(o.legendPos)) next.legendPos = o.legendPos === "r" ? undefined : o.legendPos;
  if (o.grouping && ["clustered", "stacked", "percentStacked"].includes(o.grouping)) next.grouping = o.grouping === "clustered" ? undefined : o.grouping;
  if (typeof o.dataLabels === "boolean") next.dataLabels = o.dataLabels || undefined;
  if (o.dataLabelPos && ["outEnd", "ctr", "inEnd"].includes(o.dataLabelPos)) next.dataLabelPos = o.dataLabelPos;
  if ("axisTitleX" in o) next.axisTitleX = o.axisTitleX == null || o.axisTitleX === "" ? undefined : String(o.axisTitleX);
  if ("axisTitleY" in o) next.axisTitleY = o.axisTitleY == null || o.axisTitleY === "" ? undefined : String(o.axisTitleY);
  if (typeof o.hideAxisX === "boolean") next.hideAxisX = o.hideAxisX || undefined;
  if (typeof o.hideAxisY === "boolean") next.hideAxisY = o.hideAxisY || undefined;
  if (typeof o.hideGridlines === "boolean") next.hideGridlines = o.hideGridlines || undefined;
  if ("gridColor" in o) next.gridColor = o.gridColor == null ? undefined : ref(o.gridColor);
  if (typeof o.markers === "boolean") next.marker = o.markers;
  if (typeof o.markerSizePt === "number" && o.markerSizePt > 0) next.markerSizePt = o.markerSizePt;
  if (typeof o.smooth === "boolean") next.smooth = o.smooth;
  if (o.radarStyle && ["standard", "marker", "filled"].includes(o.radarStyle)) next.radarStyle = o.radarStyle;
  if (typeof o.labelSizePt === "number" && o.labelSizePt > 0) next.labelSizePt = o.labelSizePt;
  if (Array.isArray(o.seriesColors)) {
    const sc = o.seriesColors;
    next.series = next.series.map((se, i) => { if (i >= sc.length) return se; const v = sc[i]; return { ...se, color: v == null ? undefined : ref(v) }; });
  }
  if (Array.isArray(o.sliceColors)) {
    const sc = o.sliceColors;
    next.pointColors = next.categories.map((_, i) => { const v = i < sc.length ? sc[i] : null; return v == null ? null : (ref(v) ?? null); });
  }
  if ("chartFill" in o) next.chartFill = fillOf(o.chartFill);
  if ("plotFill" in o) next.plotFill = fillOf(o.plotFill);
  return next;
}

/**
 * Host-enforced export gate — runs before any export produces bytes, on every
 * path (UI, keyboard, and pe:invoke from the console). Throws when blocked; the
 * bridge turns the throw into pe:result { ok:false, error }. The permission is
 * read from a module-scoped value the page console can't reach.
 */
async function assertExportAllowed(format: "pdf" | "png" | "zip"): Promise<void> {
  if (!permissions.export) throw new Error("export disabled");
  if (exportAuthUrl) {
    let res: Response;
    try {
      res = await fetch(exportAuthUrl, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "export", format, title: store.pres.title, slideCount: store.pres.slides.length }),
      });
    } catch {
      throw new Error("export authorization request failed");
    }
    if (!res.ok) throw new Error("export not authorized");
  }
}

// ----------------------------- the API ---------------------------------------

export interface PresentationEditorApi {
  readonly version: string;
  /** EMU per CSS px (9525) — for converting the EMU geometry fields. */
  readonly EMU_PER_PX: number;
  readonly EMU_PER_INCH: number;
  readonly EMU_PER_PT: number;

  // —— read ——
  getDocument(): DocumentInfo;
  getSlides(): SlideInfo[];
  getActiveSlide(): SlideInfo;
  getSlide(index: number): SlideInfo | null;
  getSelection(): SelectionInfo;
  getElement(id: string): ElementInfo | null;
  getElements(): ElementInfo[];
  /** The resolved host UI config (which editor chrome is shown). */
  getConfig(): UIConfig;
  /** The resolved host permissions the editor enforces (e.g. `{ export }`). */
  getPermissions(): Permissions;
  /**
   * Resolves once a document has loaded AND its slides have painted. If a
   * document is already loaded, resolves immediately; otherwise waits for the
   * next one. Cross-origin: callable over pe:invoke (replies pe:result when ready).
   */
  whenReady(): Promise<ReadyInfo>;

  // —— write: element properties (active slide; undoable; marks the doc dirty) ——
  selectSlide(index: number): void;
  selectElement(idOrIds: string | string[]): void;
  clearSelection(): void;
  setText(id: string, text: string): boolean;
  setElementProperties(id: string, props: Partial<{ x: number; y: number; w: number; h: number; rot: number; flipH: boolean; flipV: boolean }>): boolean;
  setFillColor(id: string, hex: string): boolean;
  setImage(id: string, urlOrDataUrl: string): Promise<boolean>;
  deleteElement(id: string): boolean;
  undo(): void;
  redo(): void;

  // —— write: slides & document ——
  addSlide(opts?: { layout?: LayoutKind; index?: number }): number;     // returns the new slide index
  duplicateSlide(index?: number): number;                               // returns the new slide index
  deleteSlide(index?: number): void;
  moveSlide(from: number, to: number): void;
  setDocumentTitle(title: string): void;
  applyTheme(palette: Partial<Record<SchemeSlot, string>>): void;       // recolor theme scheme slots
  setSlideBackgroundColor(hex: string): boolean;                        // active slide

  // —— write: insert new elements (active slide; each returns the new element id) ——
  insertText(opts?: InsertTextOpts): string;
  insertShape(geom: PresetGeom, opts?: InsertShapeOpts): string;
  insertImage(urlOrDataUrl: string, opts?: InsertBox): Promise<string>;
  insertChart(chartType: ChartKind, opts?: InsertChartOpts): string;
  insertTable(rows: number, cols: number, opts?: InsertBox): string;

  // —— write: element content & style ——
  setTableCell(id: string, row: number, col: number, text: string): boolean;
  setChartData(id: string, data: ChartData): boolean;
  setChartOptions(id: string, options: ChartOptions): boolean;
  setParagraphStyle(id: string, style: ParagraphStyle): boolean;
  setTextStyle(id: string, style: TextStyle): boolean;
  reorderElement(id: string, dir: "front" | "back" | "forward" | "backward"): boolean;

  // —— export (renders slides to images; the renderer & jsPDF load on demand) ——
  exportSlidePNG(index?: number, opts?: { scale?: number }): Promise<Blob>;   // default: active slide
  exportPDF(opts?: { scale?: number }): Promise<Blob>;                        // whole deck, one slide per page
  exportPNGZip(opts?: { scale?: number }): Promise<Blob>;                     // every slide as a PNG, zipped

  // —— events ——
  on(event: ApiEvent, handler: (payload: unknown) => void): () => void;
}

/** Payload reported when a document has loaded and painted. */
export interface ReadyInfo {
  title: string;
  slideCount: number;
  slideWidth: number;   // EMU
  slideHeight: number;  // EMU
  warnings: number;     // count of approximate-import warnings
}

// document-ready state, fulfilled by the app after a load has painted
let readyCount = 0;
let lastReady: ReadyInfo | null = null;
const readyWaiters: ((info: ReadyInfo) => void)[] = [];

/**
 * Called by the app once a freshly loaded document has actually painted.
 * Resolves any pending whenReady() promises. Not part of the public surface.
 */
export function signalDocumentReady(info: ReadyInfo): void {
  readyCount++;
  lastReady = info;
  readyWaiters.splice(0).forEach(fn => { try { fn(info); } catch { /* host handler threw */ } });
}

function buildApi(): PresentationEditorApi {
  const listeners: Record<ApiEvent, Set<(p: unknown) => void>> = {
    selection: new Set(), slide: new Set(), dirty: new Set(), change: new Set(),
  };
  const fire = (e: ApiEvent, payload: unknown) => listeners[e].forEach(fn => { try { fn(payload); } catch { /* host handler threw */ } });

  // diff store changes into typed events
  let last = { slide: store.getState().selection.slideIndex, ids: "", dirty: store.getState().dirty, pres: store.pres };
  store.subscribe(() => {
    const st = store.getState();
    const ids = st.selection.shapeIds.join(",");
    const slideChanged = st.selection.slideIndex !== last.slide;
    if (slideChanged) fire("slide", api.getActiveSlide());
    if (ids !== last.ids || slideChanged) fire("selection", api.getSelection());
    if (st.dirty !== last.dirty) fire("dirty", st.dirty);
    if (st.pres !== last.pres) fire("change", undefined);
    last = { slide: st.selection.slideIndex, ids, dirty: st.dirty, pres: st.pres };
  });

  const palette = (): Record<string, string> => {
    const t = store.pres.theme;
    const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"] as const;
    return Object.fromEntries(slots.map(s => [s, "#" + t[s]]));
  };

  const api: PresentationEditorApi = {
    version: typeof __APP_VERSION__ !== "undefined" ? __APP_VERSION__ : "dev",
    EMU_PER_PX, EMU_PER_INCH, EMU_PER_PT,

    getDocument() {
      const st = store.getState();
      return {
        title: store.pres.title,
        slideCount: store.pres.slides.length,
        slideWidth: store.pres.slideWidth,
        slideHeight: store.pres.slideHeight,
        activeSlideIndex: st.selection.slideIndex,
        dirty: st.dirty, canUndo: st.canUndo, canRedo: st.canRedo,
        palette: palette(),
      };
    },
    getSlides() { return store.pres.slides.map((_, i) => describeSlide(i)); },
    getActiveSlide() { return describeSlide(store.getState().selection.slideIndex, true); },
    getSlide(index) {
      return index >= 0 && index < store.pres.slides.length ? describeSlide(index, true) : null;
    },
    getSelection() {
      const sel = store.getState().selection;
      const shapes = store.selectedShapes;
      return { slideIndex: sel.slideIndex, ids: sel.shapeIds, elements: shapes.map(describeElement) };
    },
    getElement(id) { const s = findOnActive(id); return s ? describeElement(s) : null; },
    getElements() { return activeShapes().map(describeElement); },
    getConfig() { return { ...uiConfig, tabs: { ...uiConfig.tabs } }; },
    getPermissions() { return { ...permissions }; },
    whenReady() {
      return readyCount > 0 && lastReady
        ? Promise.resolve(lastReady)
        : new Promise<ReadyInfo>(resolve => readyWaiters.push(resolve));
    },

    selectSlide(index) { store.selectSlide(index); },
    selectElement(idOrIds) {
      const ids = (Array.isArray(idOrIds) ? idOrIds : [idOrIds]).filter(id => findOnActive(id));
      store.selectShapes(ids);
    },
    clearSelection() { store.selectShapes([]); },

    setText(id, text) {
      const s = findOnActive(id);
      if (!s || s.kind !== "sp") return false;
      const tmpl = templateRun(s.text);
      const paragraphs = String(text).split("\n").map(line => ({
        runs: [{ ...tmpl, text: line }], align: "l" as TextAlign, bullet: "none" as const, level: 0,
      }));
      const body = s.text ?? EMPTY_BODY();
      store.updateShapes([id], sh => sh.kind === "sp" ? { ...sh, text: { ...body, paragraphs } } : sh);
      return true;
    },

    setElementProperties(id, props) {
      const s = findOnActive(id);
      if (!s) return false;
      store.updateShapes([id], sh => ({
        ...sh,
        x: props.x ?? sh.x,
        y: props.y ?? sh.y,
        w: props.w !== undefined ? Math.max(1, Math.round(props.w)) : sh.w,
        h: props.h !== undefined ? Math.max(1, Math.round(props.h)) : sh.h,
        rot: props.rot ?? sh.rot,
        flipH: props.flipH ?? sh.flipH,
        flipV: props.flipV ?? sh.flipV,
      }));
      return true;
    },

    setFillColor(id, hex) {
      const s = findOnActive(id);
      if (!s || s.kind !== "sp") return false;
      const h = normalizeHex(hex);
      if (!h) return false;
      store.updateShapes([id], sh => sh.kind === "sp" ? { ...sh, fill: { kind: "solid", color: { kind: "srgb", hex: h } } } : sh);
      return true;
    },

    async setImage(id, urlOrDataUrl) {
      const s = findOnActive(id);
      if (!s || s.kind !== "pic") return false;
      const { mediaId } = await loadImageFromUrl(urlOrDataUrl);
      store.replacePicMedia(id, mediaId);
      return true;
    },

    deleteElement(id) {
      if (!findOnActive(id)) return false;
      store.selectShapes([id]);
      store.deleteSelectedShapes();
      return true;
    },

    undo() { store.undo(); },
    redo() { store.redo(); },

    // —— slides & document ——
    addSlide(opts = {}) {
      // store.addSlide inserts AFTER `after`; map a desired index to after = index-1
      const after = opts.index !== undefined ? opts.index - 1 : undefined;
      store.addSlide(opts.layout ?? "titleContent", after);
      return store.getState().selection.slideIndex;
    },
    duplicateSlide(index) { store.duplicateSlide(index); return store.getState().selection.slideIndex; },
    deleteSlide(index) { store.deleteSlide(index); },
    moveSlide(from, to) { store.moveSlide(from, to); },
    setDocumentTitle(title) { store.commit({ ...store.pres, title: String(title) }); },
    applyTheme(palette) {
      const theme = { ...store.pres.theme };
      for (const [slot, hex] of Object.entries(palette)) {
        const h = hex && normalizeHex(hex);
        if (h && slot in theme) (theme as unknown as Record<string, string>)[slot] = h;
      }
      store.commit({ ...store.pres, theme });
    },
    setSlideBackgroundColor(hex) {
      const h = normalizeHex(hex);
      if (!h) return false;
      store.updateSlideBg({ kind: "solid", color: { kind: "srgb", hex: h } });
      return true;
    },

    // —— insert new elements ——
    insertText(opts = {}) {
      const b = placeBox(opts, 4_000_000, 1_000_000);
      const s = makeTextBox(b.x, b.y, b.w, b.h, opts.text ?? "");
      store.addShape(s);
      return s.id;
    },
    insertShape(geom, opts = {}) {
      const b = placeBox(opts, 2_000_000, 2_000_000);
      const s = makeShape(geom, b.x, b.y, b.w, b.h);
      const h = opts.fillColor && normalizeHex(opts.fillColor);
      if (h) s.fill = { kind: "solid", color: { kind: "srgb", hex: h } };
      store.addShape(s);
      return s.id;
    },
    async insertImage(urlOrDataUrl, opts = {}) {
      const { mediaId, natW, natH } = await loadImageFromUrl(urlOrDataUrl);
      let w = opts.w ?? natW * EMU_PER_PX, h = opts.h ?? natH * EMU_PER_PX;
      if (opts.w === undefined && opts.h === undefined) {
        const k = Math.min(1, store.pres.slideWidth * 0.5 / w, store.pres.slideHeight * 0.5 / h);
        w *= k; h *= k;
      }
      const b = placeBox({ ...opts, w, h }, w, h);
      const s: PicShape = { kind: "pic", id: nextId("pic"), name: "Image", mediaId, x: b.x, y: b.y, w: b.w, h: b.h, rot: 0 };
      store.addShape(s);
      return s.id;
    },
    insertChart(chartType, opts = {}) {
      // hard guard: an invalid kind would serialize an empty plot element and make the .pptx unreadable.
      if (!CHART_KINDS.includes(chartType)) throw new Error(`insertChart: invalid chartType "${String(chartType)}". Use one of: ${CHART_KINDS.join(", ")}`);
      const b = placeBox(opts, 5_000_000, 3_000_000);
      let s = makeChart(chartType, b.x, b.y, b.w, b.h);
      if (opts.categories) s.categories = opts.categories.map(String);
      if (opts.series) s.series = opts.series.map((se, i) => ({ name: String(se.name ?? `Series ${i + 1}`), values: (se.values ?? []).map(Number) }));
      if (opts.options) s = applyChartOptions(s, opts.options);
      store.addShape(s);
      return s.id;
    },
    insertTable(rows, cols, opts = {}) {
      const r = Math.max(1, Math.round(rows)), c = Math.max(1, Math.round(cols));
      const b = placeBox(opts, 5_000_000, r * 600_000);
      const s = makeTable(r, c, b.x, b.y, b.w, b.h);
      store.addShape(s);
      return s.id;
    },

    // —— element content & style ——
    setTableCell(id, row, col, text) {
      const s = findOnActive(id);
      if (!s || s.kind !== "table") return false;
      if (row < 0 || col < 0 || row >= s.cells.length || col >= (s.cells[0]?.length ?? 0)) return false;
      store.updateShapes([id], sh => {
        if (sh.kind !== "table") return sh;
        const cells = sh.cells.map(rw => rw.map(cl => ({ ...cl })));
        const cell = cells[row][col];
        const tmpl = templateRun(cell.text);
        cell.text = { ...cell.text, paragraphs: String(text).split("\n").map(line => ({ runs: [{ ...tmpl, text: line }], align: "l" as TextAlign, bullet: "none" as const, level: 0 })) };
        return { ...sh, cells };
      });
      return true;
    },
    setChartData(id, data) {
      const s = findOnActive(id);
      if (!s || s.kind !== "chart") return false;
      store.updateShapes([id], sh => {
        if (sh.kind !== "chart") return sh;
        const next = { ...sh };
        if (data.categories) next.categories = data.categories.map(String);
        if (data.series) next.series = data.series.map((se, i) => ({ name: String(se.name ?? sh.series[i]?.name ?? `Series ${i + 1}`), values: (se.values ?? []).map(Number), color: sh.series[i]?.color }));
        return next;
      });
      return true;
    },
    setChartOptions(id, options) {
      const s = findOnActive(id);
      if (!s || s.kind !== "chart") return false;
      store.updateShapes([id], sh => (sh.kind === "chart" ? applyChartOptions(sh, options ?? {}) : sh));
      return true;
    },
    setParagraphStyle(id, style) {
      if (!findOnActive(id)) return false;
      store.formatParagraphs(p => ({ ...p, align: style.align ?? p.align, bullet: style.bullet ?? p.bullet, level: style.level ?? p.level }), [id]);
      return true;
    },
    setTextStyle(id, style) {
      if (!findOnActive(id)) return false;
      const colorHex = style.color ? normalizeHex(style.color) : null;
      store.formatRuns(r => ({
        ...r,
        bold: style.bold ?? r.bold, italic: style.italic ?? r.italic,
        underline: style.underline ?? r.underline, strike: style.strike ?? r.strike,
        sizePt: style.sizePt ?? r.sizePt, font: style.font ?? r.font,
        color: colorHex ? { kind: "srgb", hex: colorHex } : r.color,
      }), [id]);
      return true;
    },
    reorderElement(id, dir) {
      if (!findOnActive(id)) return false;
      store.reorderShape(id, dir);
      return true;
    },

    // —— export (gated by permissions; export.tsx + jsPDF load only when allowed) ——
    async exportSlidePNG(index, opts) {
      await assertExportAllowed("png");
      const i = index ?? store.getState().selection.slideIndex;
      const slide = store.pres.slides[i];
      if (!slide) throw new Error(`slide index ${i} out of range`);
      const { slideToPngBlob } = await import("./export");
      return slideToPngBlob(store.pres, slide, store.media, opts);
    },
    async exportPDF(opts) {
      await assertExportAllowed("pdf");
      const { exportPdfBlob } = await import("./export");
      return exportPdfBlob(store.pres, store.media, opts);
    },
    async exportPNGZip(opts) {
      await assertExportAllowed("zip");
      const { exportPngZipBlob } = await import("./export");
      return exportPngZipBlob(store.pres, store.media, opts);
    },

    on(event, handler) {
      listeners[event].add(handler);
      return () => listeners[event].delete(handler);
    },
  };
  return api;
}

/** The singleton public API, backed by the editor's singleton store. */
export const editorApi: PresentationEditorApi = buildApi();

/** Method names the postMessage bridge is allowed to invoke (read + write). */
export const API_METHODS = [
  "getDocument", "getSlides", "getActiveSlide", "getSlide", "getSelection", "getElement", "getElements", "getConfig", "getPermissions", "whenReady",
  "selectSlide", "selectElement", "clearSelection",
  "setText", "setElementProperties", "setFillColor", "setImage", "deleteElement", "undo", "redo",
  "addSlide", "duplicateSlide", "deleteSlide", "moveSlide", "setDocumentTitle", "applyTheme", "setSlideBackgroundColor",
  "insertText", "insertShape", "insertImage", "insertChart", "insertTable",
  "setTableCell", "setChartData", "setChartOptions", "setParagraphStyle", "setTextStyle", "reorderElement",
  "exportSlidePNG", "exportPDF", "exportPNGZip",
] as const;

/** Install the API as a global so same-origin hosts can reach it. */
export function installApi(win: Window) {
  (win as unknown as { presentationEditor: PresentationEditorApi }).presentationEditor = editorApi;
}
