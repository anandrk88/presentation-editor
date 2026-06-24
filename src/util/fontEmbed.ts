/**
 * Collect the bundled font faces a presentation actually uses and fetch their
 * bytes, so the .pptx writer can embed them. Browser-only (uses fetch). The font
 * files are self-hosted next to fonts.css (served at <base>/fonts/<file>), so we
 * resolve them relative to document.baseURI exactly like the <link> does.
 *
 * Only fonts we bundle are embedded — system fonts (Arial, Calibri, …) are present
 * everywhere and not ours to ship. Faces whose OS/2 fsType forbids embedding are
 * skipped. Embedding is opt-in (the explicit "Export .pptx") — the lean save path
 * (pe:save → your storage) never calls this.
 */
import type { Presentation } from "../model/types";
import { resolveFontName } from "../model/defaults";
import { FONT_FILES } from "../fonts/manifest";
import type { EmbeddedFont } from "../ooxml/write";

/** Every typeface the deck references, resolved (+mj-lt/+mn-lt → theme fonts). */
function usedTypefaces(pres: Presentation): Set<string> {
  const theme = pres.theme;
  const out = new Set<string>();
  const add = (font?: string) => { if (font) out.add(resolveFontName(font, theme)); };
  add(theme.majorFont);
  add(theme.minorFont);
  const scan = (body?: { paragraphs: { runs: { font: string }[] }[] }) =>
    body?.paragraphs.forEach(p => p.runs.forEach(r => add(r.font)));
  for (const slide of pres.slides) {
    for (const s of slide.shapes) {
      if (s.kind === "sp") scan(s.text);
      else if (s.kind === "table") s.cells.forEach(row => row.forEach(c => scan(c.text)));
      else if (s.kind === "chart" && s.partStyles) Object.values(s.partStyles).forEach(st => add(st?.font));
    }
  }
  return out;
}

/** fsType lives in the OS/2 table; embedding is forbidden only when bit 1 (0x0002) is set. */
function embeddingAllowed(bytes: Uint8Array): boolean {
  try {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (dv.byteLength < 12) return true;
    const numTables = dv.getUint16(4);
    for (let i = 0, off = 12; i < numTables; i++, off += 16) {
      if (off + 16 > dv.byteLength) break;
      const tag = String.fromCharCode(dv.getUint8(off), dv.getUint8(off + 1), dv.getUint8(off + 2), dv.getUint8(off + 3));
      if (tag === "OS/2") {
        const tableOff = dv.getUint32(off + 8);
        if (tableOff + 10 > dv.byteLength) return true;
        return (dv.getUint16(tableOff + 8) & 0x0002) === 0;   // 0x0002 = Restricted License embedding
      }
    }
  } catch { /* unparseable header → allow (our bundled set is OFL/Apache) */ }
  return true;
}

const slotOf = (weight: number, italic: boolean): "regular" | "bold" | "italic" | "boldItalic" =>
  italic ? (weight >= 600 ? "boldItalic" : "italic") : (weight >= 600 ? "bold" : "regular");

/** Gather embeddable faces for every bundled typeface the deck uses. */
export async function gatherEmbeddedFonts(pres: Presentation): Promise<EmbeddedFont[]> {
  const families = usedTypefaces(pres);
  const out: EmbeddedFont[] = [];
  for (const fam of families) {
    const faces = FONT_FILES.filter(f => f.family === fam);
    if (!faces.length) continue;   // not a bundled font → leave it to PowerPoint's substitution
    const ef: EmbeddedFont = { typeface: fam };
    for (const face of faces) {
      try {
        const res = await fetch(new URL(`fonts/${face.file}`, document.baseURI).href);
        if (!res.ok) continue;
        const bytes = new Uint8Array(await res.arrayBuffer());
        if (bytes.length && embeddingAllowed(bytes)) ef[slotOf(face.weight, face.italic)] = bytes;
      } catch { /* skip a face that fails to load */ }
    }
    if (ef.regular || ef.bold || ef.italic || ef.boldItalic) out.push(ef);
  }
  return out;
}

/** Total bytes that would be embedded — for a "+N MB fonts" UX hint. */
export function embeddedFontsSize(fonts: EmbeddedFont[]): number {
  return fonts.reduce((n, f) =>
    n + (f.regular?.length ?? 0) + (f.bold?.length ?? 0) + (f.italic?.length ?? 0) + (f.boldItalic?.length ?? 0), 0);
}
