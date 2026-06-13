import type { PresetGeom } from "../model/types";
import { hasPreset, presetOutline } from "./presetGeom";

/**
 * Geometry facade. All preset outlines now come from the ECMA-376 evaluator
 * (src/render/presetGeom.ts) — the same definition tables PowerPoint uses —
 * so every preset renders correctly at any aspect ratio.
 */
export function presetPath(geom: PresetGeom, w: number, h: number, adj?: Record<string, number>): string {
  if (hasPreset(geom)) return presetOutline(geom, w, h, adj);
  return `M0 0H${w}V${h}H0Z`;
}

/** Connector-style geometries: stroked, never filled, no text body by default. */
const LINE_GEOMS = new Set([
  "line", "lineInv", "straightConnector1",
  "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5",
  "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5",
]);

export function isLineGeom(g: PresetGeom): boolean {
  return LINE_GEOMS.has(g);
}

export interface ShapeCategory {
  label: string;
  shapes: PresetGeom[];
}

/** Categorized gallery mirroring PowerPoint's insert-shape panel. */
export const SHAPE_CATEGORIES: ShapeCategory[] = [
  {
    label: "Basic shapes",
    shapes: [
      "rect", "ellipse", "triangle", "rtTriangle", "parallelogram", "trapezoid",
      "diamond", "pentagon", "hexagon", "heptagon", "octagon", "decagon", "dodecagon",
      "pie", "chord", "teardrop", "frame", "halfFrame", "corner", "diagStripe",
      "plus", "plaque", "can", "cube", "bevel", "donut", "noSmoking", "blockArc",
      "foldedCorner", "smileyFace", "heart", "lightningBolt", "sun", "moon", "cloud",
      "arc", "bracketPair", "bracePair", "leftBracket", "rightBracket", "leftBrace", "rightBrace",
    ],
  },
  {
    label: "Figured arrows",
    shapes: [
      "rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow",
      "quadArrow", "leftRightUpArrow", "bentArrow", "uturnArrow", "leftUpArrow", "bentUpArrow",
      "curvedRightArrow", "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow",
      "stripedRightArrow", "notchedRightArrow", "homePlate", "chevron",
      "rightArrowCallout", "downArrowCallout", "leftArrowCallout", "upArrowCallout",
      "leftRightArrowCallout", "quadArrowCallout", "circularArrow",
    ],
  },
  {
    label: "Math",
    shapes: ["mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual"],
  },
  {
    label: "Charts",
    shapes: [
      "flowChartProcess", "flowChartAlternateProcess", "flowChartDecision", "flowChartInputOutput",
      "flowChartPredefinedProcess", "flowChartInternalStorage", "flowChartDocument", "flowChartMultidocument",
      "flowChartTerminator", "flowChartPreparation", "flowChartManualInput", "flowChartManualOperation",
      "flowChartConnector", "flowChartOffpageConnector", "flowChartPunchedCard", "flowChartPunchedTape",
      "flowChartSummingJunction", "flowChartOr", "flowChartCollate", "flowChartSort",
      "flowChartExtract", "flowChartMerge", "flowChartOnlineStorage", "flowChartDelay",
      "flowChartMagneticTape", "flowChartMagneticDisk", "flowChartMagneticDrum", "flowChartDisplay",
    ],
  },
  {
    label: "Stars & ribbons",
    shapes: [
      "irregularSeal1", "irregularSeal2", "star4", "star5", "star6", "star7", "star8",
      "star10", "star12", "star16", "star24", "star32",
      "ribbon", "ribbon2", "ellipseRibbon", "ellipseRibbon2",
      "verticalScroll", "horizontalScroll", "wave", "doubleWave",
    ],
  },
  {
    label: "Callouts",
    shapes: [
      "wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout",
      "borderCallout1", "borderCallout2", "borderCallout3",
      "accentCallout1", "accentCallout2", "accentCallout3",
      "callout1", "callout2", "callout3",
      "accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3",
    ],
  },
  {
    label: "Buttons",
    shapes: [
      "actionButtonBackPrevious", "actionButtonForwardNext", "actionButtonBeginning", "actionButtonEnd",
      "actionButtonHome", "actionButtonInformation", "actionButtonReturn", "actionButtonMovie",
      "actionButtonDocument", "actionButtonSound", "actionButtonHelp", "actionButtonBlank",
    ],
  },
  {
    label: "Rectangles",
    shapes: [
      "rect", "roundRect", "snip1Rect", "snip2SameRect", "snip2DiagRect",
      "snipRoundRect", "round1Rect", "round2SameRect", "round2DiagRect",
    ],
  },
  {
    label: "Lines",
    shapes: ["line", "straightConnector1", "bentConnector3", "curvedConnector3"],
  },
];

const SPECIAL_LABELS: Record<string, string> = {
  rect: "Rectangle", roundRect: "Rounded Rectangle", ellipse: "Ellipse", rtTriangle: "Right Triangle",
  homePlate: "Pentagon Arrow", line: "Line", straightConnector1: "Line", bentConnector3: "Elbow Connector",
  curvedConnector3: "Curved Connector", noSmoking: "No Symbol", uturnArrow: "U-Turn Arrow",
};

export function geomLabel(g: PresetGeom): string {
  if (SPECIAL_LABELS[g]) return SPECIAL_LABELS[g];
  return g
    .replace(/([a-z])([A-Z0-9])/g, "$1 $2")
    .replace(/^./, c => c.toUpperCase())
    .replace(/Flow Chart/, "")
    .replace(/Action Button/, "")
    .trim() || g;
}

/** Compact list for the inline Home-toolbar strip. */
export const SHAPE_GALLERY: { geom: PresetGeom; label: string }[] = [
  "rect", "roundRect", "ellipse", "triangle", "rtTriangle", "diamond", "parallelogram",
  "trapezoid", "pentagon", "hexagon", "chevron", "rightArrow", "leftArrow", "upArrow",
  "downArrow", "star5", "heart", "line",
].map(geom => ({ geom, label: geomLabel(geom) }));
