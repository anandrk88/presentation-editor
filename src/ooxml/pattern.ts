import JSZip from "jszip";
import type { MediaItem, SlideModel } from "../model/types";
import { parsePptx } from "./parse";
import { esc } from "./write";

/**
 * SlideBazaar-style pattern import: a pattern JSON carries raw PresentationML
 * fragments (elements[].xml) with {{content.*}} placeholder tokens. We fill the
 * tokens, wrap the fragment in a minimal one-slide pptx package in memory, and
 * run it through the regular OOXML parser — so patterns get the exact same
 * fidelity (custGeom, gradients, scheme colors, groups) as opened files.
 */

export interface PatternJson {
  id?: string;
  name?: string;
  description?: string;
  ai_hint?: string;
  content_schema?: Record<string, { type?: string; max?: number; required?: boolean; hint?: string }>;
  elements?: { type: string; xml?: string; _note?: string }[];
}

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;

const SAMPLES: [RegExp, string][] = [
  [/year/i, String(new Date().getFullYear())],
  [/company|brand|org/i, "Acme Analytics"],
  [/tagline/i, "Charting sustainable growth across every market"],
  [/subtitle/i, "Annual performance and market insights"],
  [/title/i, "Growth Report"],
  [/date/i, "June 2026"],
  [/email/i, "hello@acme.com"],
  [/url|web/i, "www.acme.com"],
];

function clampWords(s: string, max?: number): string {
  if (!max || s.length <= max) return s;
  const words = s.split(/\s+/);
  let out = "";
  for (const w of words) {
    const next = out ? `${out} ${w}` : w;
    if (next.length > max) break;
    out = next;
  }
  return out || s.slice(0, max);
}

export function sampleValue(key: string, max?: number): string {
  for (const [re, v] of SAMPLES) if (re.test(key)) return clampWords(v, max);
  return clampWords("Sample text", max);
}

export function fillPlaceholders(xml: string, values: Record<string, string>): string {
  return xml.replace(/\{\{\s*([\w.]+)\s*\}\}/g, (_, key: string) => {
    const bare = key.replace(/^content\./, "");
    const v = values[key] ?? values[bare];
    return v !== undefined ? esc(v) : "";
  });
}

export interface ImportedPattern {
  slide: SlideModel;
  media: Map<string, MediaItem>;
  warnings: string[];
  values: Record<string, string>;
}

export async function importPatternSlide(
  json: string | PatternJson,
  userValues?: Record<string, string>,
): Promise<ImportedPattern> {
  let pat: PatternJson;
  try {
    pat = typeof json === "string" ? (JSON.parse(json) as PatternJson) : json;
  } catch (e) {
    throw new Error("Not valid JSON: " + (e as Error).message);
  }
  const elements = (pat.elements ?? []).filter(e => e.type === "ooxml" && e.xml);
  if (!elements.length) throw new Error("Pattern has no elements of type 'ooxml'");

  // placeholder values: schema-driven samples, overridable by caller
  const values: Record<string, string> = {};
  for (const [key, spec] of Object.entries(pat.content_schema ?? {})) {
    if (key === "__proto__" || key === "constructor" || key === "prototype") continue;
    values[key] = userValues?.[key] !== undefined
      ? clampWords(userValues[key], spec?.max)
      : sampleValue(key, spec?.max);
  }
  Object.assign(values, userValues ?? {});

  const fragment = elements.map(e => fillPlaceholders(e.xml!, values)).join("");

  const slideXml = XML_DECL +
    `<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    fragment +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

  // minimal in-memory package — just enough for parsePptx to walk the rels
  const zip = new JSZip();
  zip.file("_rels/.rels", XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file("ppt/presentation.xml", XML_DECL +
    `<p:presentation xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>` +
    `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file("ppt/_rels/presentation.xml.rels", XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file("ppt/slides/slide1.xml", slideXml);

  const bytes = await zip.generateAsync({ type: "uint8array" });
  const { pres, media, warnings } = await parsePptx(bytes, pat.name ?? pat.id ?? "Pattern");
  return { slide: pres.slides[0], media, warnings, values };
}
