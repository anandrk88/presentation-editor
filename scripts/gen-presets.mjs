// Compiles ECMA-376 presetShapeDefinitions.xml (vendored from LibreOffice's copy
// of the spec) into src/render/presetDefs.json: every PowerPoint preset geometry
// as adjust values, guide formulas and path commands for the runtime evaluator.
import { DOMParser } from "@xmldom/xmldom";
import { readFileSync, writeFileSync } from "fs";
import { dirname, resolve } from "path";
import { fileURLToPath } from "url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const xml = readFileSync(resolve(root, "scripts/data/presetShapeDefinitions.xml"), "utf8");
const doc = new DOMParser().parseFromString(xml, "application/xml");

const kids = (el, name) => {
  const out = [];
  for (let i = 0; i < el.childNodes.length; i++) {
    const n = el.childNodes[i];
    if (n.nodeType === 1 && (!name || n.localName === name)) out.push(n);
  }
  return out;
};
const kid = (el, name) => kids(el, name)[0] ?? null;

const defs = {};
let presetCount = 0;
for (const preset of kids(doc.documentElement)) {
  const name = preset.localName;
  const entry = {};

  const av = kids(kid(preset, "avLst") ?? preset, "gd")
    .filter(g => g.parentNode.localName === "avLst")
    .map(g => [g.getAttribute("name"), g.getAttribute("fmla")]);
  if (av.length) entry.av = av;

  const gd = kid(preset, "gdLst")
    ? kids(kid(preset, "gdLst"), "gd").map(g => [g.getAttribute("name"), g.getAttribute("fmla")])
    : [];
  if (gd.length) entry.gd = gd;

  const rect = kid(preset, "rect");
  if (rect) entry.rect = ["l", "t", "r", "b"].map(a => rect.getAttribute(a));

  const pathLst = kid(preset, "pathLst");
  entry.paths = kids(pathLst, "path").map(p => {
    const out = {};
    if (p.getAttribute("w")) out.w = Number(p.getAttribute("w"));
    if (p.getAttribute("h")) out.h = Number(p.getAttribute("h"));
    const fill = p.getAttribute("fill");
    if (fill && fill !== "norm") out.fill = fill;
    if (p.getAttribute("stroke") === "false") out.stroke = 0;
    const cmds = [];
    for (const c of kids(p)) {
      switch (c.localName) {
        case "moveTo": {
          const pt = kid(c, "pt");
          cmds.push(["M", pt.getAttribute("x"), pt.getAttribute("y")]);
          break;
        }
        case "lnTo": {
          const pt = kid(c, "pt");
          cmds.push(["L", pt.getAttribute("x"), pt.getAttribute("y")]);
          break;
        }
        case "arcTo":
          cmds.push(["A", c.getAttribute("wR"), c.getAttribute("hR"), c.getAttribute("stAng"), c.getAttribute("swAng")]);
          break;
        case "cubicBezTo": {
          const pts = kids(c, "pt");
          cmds.push(["C", ...pts.flatMap(pt => [pt.getAttribute("x"), pt.getAttribute("y")])]);
          break;
        }
        case "quadBezTo": {
          const pts = kids(c, "pt");
          cmds.push(["Q", ...pts.flatMap(pt => [pt.getAttribute("x"), pt.getAttribute("y")])]);
          break;
        }
        case "close":
          cmds.push(["Z"]);
          break;
      }
    }
    out.cmds = cmds;
    return out;
  });

  defs[name] = entry;
  presetCount++;
}

const outPath = resolve(root, "src/render/presetDefs.json");
writeFileSync(outPath, JSON.stringify(defs));
console.log(`wrote ${outPath}: ${presetCount} presets, ${(JSON.stringify(defs).length / 1024).toFixed(0)} KB`);
