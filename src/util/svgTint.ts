/**
 * Single-color recoloring for SVG graphics — same semantics as PowerPoint's
 * "Graphics Fill": every painted fill/stroke becomes the chosen color.
 * Pure string transforms (used by the renderer in the browser and the
 * pptx writer in node).
 */

export function tintSvgText(text: string, hex: string): string {
  // Office theme-bound icons: PowerPoint re-binds any class named
  // MsftOfcThm_<slot>_Fill/Stroke to the LIVE theme color at render time,
  // overriding the literal CSS — rename the magic prefix or the explicit
  // tint silently reverts (e.g. white shows as accent1).
  let out = text.replace(/MsftOfcThm_/g, "TintedThm_");
  out = out.replace(/fill="(?!none)[^"]*"/g, `fill="${hex}"`);
  out = out.replace(/stroke="(?!none)[^"]*"/g, `stroke="${hex}"`);
  out = out.replace(/fill:\s*(?!none)[#a-zA-Z0-9(),.\s%]+?([;"}])/g, `fill:${hex}$1`);
  out = out.replace(/stroke:\s*(?!none)[#a-zA-Z0-9(),.\s%]+?([;"}])/g, `stroke:${hex}$1`);
  // shapes with no fill attribute default to black — recolor those via a root fill
  if (!/<svg[^>]*\sfill=/.test(out)) {
    out = out.replace(/<svg/, `<svg fill="${hex}"`);
  }
  return out;
}

/** Flatten any resolved CSS color ("#RRGGBB" or "rgba(...)") to #RRGGBB. */
export function cssColorToHex(css: string): string {
  if (css.startsWith("#")) return css.length === 4
    ? "#" + [...css.slice(1)].map(c => c + c).join("")
    : css.slice(0, 7);
  const m = css.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!m) return "#000000";
  return "#" + [m[1], m[2], m[3]].map(v => parseInt(v, 10).toString(16).padStart(2, "0")).join("");
}
