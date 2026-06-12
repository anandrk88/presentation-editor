// Document model mirroring OOXML PresentationML structures.
// All geometry is stored in EMU (914400 per inch, 9525 per CSS px @96dpi).

export type EMU = number;

export interface ColorTheme {
  name: string;
  // scheme slots, hex without '#'
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  folHlink: string;
  majorFont: string;
  minorFont: string;
}

export type SchemeSlot =
  | "dk1" | "lt1" | "dk2" | "lt2"
  | "accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6"
  | "hlink" | "folHlink";

/** A color is either a literal sRGB hex or a reference into the theme scheme. */
export type ColorRef =
  | { kind: "srgb"; hex: string; lumMod?: number; lumOff?: number; alpha?: number }
  | { kind: "scheme"; slot: SchemeSlot; lumMod?: number; lumOff?: number; alpha?: number };

export interface GradientStop {
  pos: number; // 0-100
  color: ColorRef;
}

export type Fill =
  | { kind: "none" }
  | { kind: "solid"; color: ColorRef }
  | { kind: "gradient"; stops: GradientStop[]; angle: number } // angle in degrees, OOXML convention (0 = left->right, cw)
  | { kind: "image"; mediaId: string; tile?: boolean }          // a:blipFill (stretch or tile)
  | { kind: "pattern"; prst: string; fg: ColorRef; bg: ColorRef }; // a:pattFill preset hatch

export interface LineProps {
  fill: Fill;          // outline color (none => no outline)
  widthPt: number;     // outline width in points
  dash?: "solid" | "dash" | "dot";
}

export type TextAlign = "l" | "ctr" | "r" | "just";
export type TextAnchor = "t" | "ctr" | "b";
export type BulletKind = "none" | "char" | "num";

export interface Run {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Baseline offset percent: +30 superscript, -25 subscript (OOXML baseline/1000). */
  baseline?: number;
  /** Text highlight color (a:highlight). */
  highlight?: ColorRef;
  sizePt: number;
  font: string;
  color: ColorRef;
}

export interface Paragraph {
  runs: Run[];
  align: TextAlign;
  bullet: BulletKind;
  level: number;          // 0-8 indent level
  lineSpacingPct?: number; // 100 = single
}

export interface TextBody {
  paragraphs: Paragraph[];
  anchor: TextAnchor;
  wrap: boolean;
  // insets in EMU (l, t, r, b) — defaults match OOXML bodyPr defaults
  insets: [EMU, EMU, EMU, EMU];
  /** Text columns (bodyPr numCol), 1 = single column. */
  columns?: number;
  /** Gap between columns in EMU (bodyPr spcCol). */
  colSpacing?: EMU;
}

/**
 * Any ECMA-376 preset geometry name (all 187 are evaluated from the spec's
 * definition tables at render time — see src/render/presetGeom.ts).
 */
export type PresetGeom = string;

export interface ShapeBase {
  id: string;
  name: string;
  x: EMU;
  y: EMU;
  w: EMU;
  h: EMU;
  rot: number; // degrees clockwise
  flipH?: boolean;
  flipV?: boolean;
  /** Shapes sharing a groupId select/move/resize as one unit and export inside a p:grpSp. */
  groupId?: string;
}

/** Parsed a:custGeom — SVG path data in the path's own coordinate space. */
export interface CustPath {
  w: number; // a:path/@w (coordinate space width)
  h: number;
  d: string;
}

export interface SpShape extends ShapeBase {
  kind: "sp";
  geom: PresetGeom;
  /** Adjust values (avLst overrides) — e.g. corner radius, arrow head depth. */
  adj?: Record<string, number>;
  /** When present, overrides geom for rendering (freeform shape). */
  custPath?: CustPath;
  /** Original <a:custGeom> XML, written back verbatim on export. */
  rawGeomXml?: string;
  isTextBox?: boolean;
  fill: Fill;
  line: LineProps;
  text?: TextBody;
}

export interface PicShape extends ShapeBase {
  kind: "pic";
  mediaId: string;
  /** Recolor for SVG graphics (PowerPoint's "Graphics Fill") — baked into the svg on export. */
  svgTint?: ColorRef;
  /** Frame geometry the image is clipped to (rect default; roundRect/ellipse for rounded/circle). */
  geom?: PresetGeom;
  /** Adjust values for the frame geometry (e.g. roundRect corner radius). */
  adj?: Record<string, number>;
  /** Crop rectangle (a:srcRect): fractions 0-1 cut from each source edge. */
  srcRect?: { l: number; t: number; r: number; b: number };
}

// ---------- tables (a:tbl inside p:graphicFrame) ----------
export interface TableCell {
  text: TextBody;
  fill?: Fill;       // explicit cell fill; undefined -> style-derived (header/banding)
  gridSpan?: number;   // horizontal merge width
  rowSpan?: number;    // vertical merge height
  merged?: "h" | "v";  // covered by a spanning cell (hMerge/vMerge) — not rendered
}

export type TableStyleFamily = "light" | "medium" | "dark";

export interface TableShape extends ShapeBase {
  kind: "table";
  colW: EMU[];           // column width weights (normalized to shape width on render/export)
  rowH: EMU[];
  cells: TableCell[][];  // [row][col]
  firstRow: boolean;     // header-row styling
  bandRow: boolean;      // banded-row styling
  /** PowerPoint "Total Row" (a:tblPr lastRow). */
  totalRow?: boolean;
  /** Emphasized first/last column (a:tblPr firstCol/lastCol). */
  firstCol?: boolean;
  lastCol?: boolean;
  /** Banded columns (a:tblPr bandCol). */
  bandCol?: boolean;
  /** Built-in style family — the Light/Medium/Dark rows of PowerPoint's gallery. */
  styleFamily?: TableStyleFamily; // default "medium"
  /** Style accent; "none" = grayscale. */
  accent?: SchemeSlot | "none";   // default "accent1"
  /** Which cell borders are drawn. */
  borderMode?: "all" | "outside" | "none"; // default "all"
  /** Border color override (default depends on the style family). */
  borderColor?: ColorRef;
}

// ---------- charts (c:chartSpace part referenced from p:graphicFrame) ----------
export type ChartKind = "column" | "bar" | "line" | "pie" | "doughnut" | "area" | "scatter" | "radar";
export type ChartGrouping = "clustered" | "stacked" | "percentStacked";
export type RadarStyle = "standard" | "marker" | "filled";

export interface ChartSeries {
  name: string;
  values: number[];
  color?: ColorRef; // undefined -> accent cycle
  /** Line thickness for line/scatter/radar series (a:ln w). */
  lineWidthPt?: number;
  /** Line dash for line/scatter/radar series. */
  dash?: "solid" | "dash" | "dot";
}

export interface ChartShape extends ShapeBase {
  kind: "chart";
  chart: ChartKind;
  /** column/bar/line/area variant (default clustered / standard). */
  grouping?: ChartGrouping;
  /** point markers on line/scatter series. */
  marker?: boolean;
  /** smoothed line/scatter curves. */
  smooth?: boolean;
  /** radar presentation (standard | marker | filled). */
  radarStyle?: RadarStyle;
  /** for scatter: categories hold the X values (numeric strings). */
  categories: string[];
  series: ChartSeries[];
  title?: string;
  legend: boolean;
  /** Legend placement (c:legendPos), default right. */
  legendPos?: "r" | "b" | "t" | "l";
  /** Axis/legend/label text size (c:txPr defRPr), default ~8pt. */
  labelSizePt?: number;
  /** Value labels on points/bars/slices (c:dLbls showVal / showPercent for pies). */
  dataLabels?: boolean;
  /** Percentage error bars on each series (c:errBars), undefined = off. */
  errorBarsPct?: number;
  axisTitleX?: string;
  axisTitleY?: string;
  hideAxisX?: boolean;
  hideAxisY?: boolean;
  /** Chart area background (c:chartSpace/c:spPr). */
  chartFill?: Fill;
  /** Chart area outline (c:chartSpace/c:spPr/a:ln). */
  chartBorder?: LineProps;
  /** Plot area background (c:plotArea/c:spPr). */
  plotFill?: Fill;
  /** Plot area outline (c:plotArea/c:spPr/a:ln). */
  plotBorder?: LineProps;
  /** Point marker size in pt (c:ser/c:marker/c:size), default ~5. */
  markerSizePt?: number;
  /** Per-slice colors for pie/doughnut (c:dPt overrides); null = automatic accent. */
  pointColors?: (ColorRef | null)[];
  /** Gridline color override (c:majorGridlines/c:spPr/a:ln). */
  gridColor?: ColorRef;
  /** Hide value-axis gridlines entirely. */
  hideGridlines?: boolean;
}

export type Shape = SpShape | PicShape | TableShape | ChartShape;

export interface SlideTransition {
  type: "none" | "fade" | "push";
  /** push direction: where the new slide comes from */
  dir?: "l" | "r" | "u" | "d";
  speed: "slow" | "med" | "fast";
}

export interface SlideModel {
  id: string;
  background?: Fill; // undefined => inherit master (white)
  shapes: Shape[];
  transition?: SlideTransition;
  notes?: string;
}

export interface MediaItem {
  id: string;
  mime: string;       // image/png, image/jpeg, image/gif, image/svg+xml
  dataUrl: string;    // data: url for rendering
  bytes: Uint8Array;  // raw bytes for pptx export
  /**
   * For SVG media: rasterized PNG written as the main a:blip fallback.
   * PowerPoint only renders vectors via the asvg:svgBlip extension and
   * requires a bitmap fallback blip alongside it.
   */
  pngFallback?: Uint8Array;
}

export interface Presentation {
  slideWidth: EMU;
  slideHeight: EMU;
  slides: SlideModel[];
  theme: ColorTheme;
  title: string;
}
