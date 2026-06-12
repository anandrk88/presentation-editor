// Diagnostic entry: expose the engine pieces needed to reproduce the
// "white icon comes back accent1 in PowerPoint" report against a real deck.
export { parsePptx, setDOMParser, setXMLSerializer } from "../src/ooxml/parse";
export { buildPptx } from "../src/ooxml/write";
export { resolveColor } from "../src/model/defaults";
export { tintSvgText, cssColorToHex } from "../src/util/svgTint";
