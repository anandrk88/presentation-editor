// Repro: open the user's real deck, recolor an svg icon white, export,
// then dump exactly what PowerPoint will see (pic XML + media bytes).
import { build } from "vite";
import { fileURLToPath, pathToFileURL } from "url";
import { dirname, resolve } from "path";
import { readFileSync, writeFileSync, mkdirSync } from "fs";
import JSZip from "jszip";
import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const SRC = process.argv[2] ?? "C:/Users/anand/Downloads/Super El Niño 2.pptx";

mkdirSync(resolve(root, ".smoke"), { recursive: true });
await build({
  configFile: false,
  root,
  logLevel: "error",
  build: {
    lib: { entry: resolve(root, "scripts/tint-entry.ts"), formats: ["es"], fileName: "tint-engine" },
    outDir: resolve(root, ".smoke"),
    emptyOutDir: false,
    minify: false,
  },
});
const eng = await import(pathToFileURL(resolve(root, ".smoke", "tint-engine.js")).href + `?t=${Date.now()}`);
eng.setDOMParser(XmldomParser);
eng.setXMLSerializer(XmldomSerializer);

const buf = readFileSync(SRC);

// ---- 1) how does the ORIGINAL PowerPoint file persist its icons? ----
const orig = await JSZip.loadAsync(buf);
console.log("=== ORIGINAL FILE: pic blipFill markup per slide (svg icons) ===");
for (const name of Object.keys(orig.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort()) {
  const xml = await orig.file(name).async("text");
  const pics = xml.match(/<p:pic>[\s\S]*?<\/p:pic>/g) ?? [];
  pics.forEach((p, i) => {
    if (!p.includes("svgBlip")) return;
    const nm = p.match(/name="([^"]*)"/)?.[1];
    const blipFill = p.match(/<p:blipFill>[\s\S]*?<\/p:blipFill>/)?.[0] ?? "(none)";
    const spPrFill = p.match(/<p:spPr>[\s\S]*?<\/p:spPr>/)?.[0]?.match(/<a:solidFill>[\s\S]*?<\/a:solidFill>/)?.[0];
    console.log(`\n--- ${name} pic#${i} "${nm}"`);
    console.log("  blipFill:", blipFill.replace(/\s+/g, " ").slice(0, 600));
    if (spPrFill) console.log("  spPr solidFill:", spPrFill.replace(/\s+/g, " "));
  });
}

// ---- 2) parse with our engine, tint one icon white, export ----
const { pres, media, warnings } = await eng.parsePptx(buf, "super.pptx");
console.log("\n=== PARSE === warnings:", warnings.length);
let target = null;
for (const [si, slide] of pres.slides.entries()) {
  for (const s of slide.shapes) {
    if (s.kind === "pic" && media.get(s.mediaId)?.mime === "image/svg+xml") {
      console.log(`slide ${si + 1}: svg pic "${s.name}" media=${s.mediaId} tint=${JSON.stringify(s.svgTint)} pngFallback=${!!media.get(s.mediaId).pngFallback}`);
      if (!target) target = s;
    }
  }
}
if (!target) { console.log("NO SVG PICS FOUND"); process.exit(1); }

target.svgTint = { kind: "srgb", hex: "FFFFFF" };
console.log(`\nTinted "${target.name}" (media ${target.mediaId}) white. Exporting…`);

const zip = await eng.buildPptx(pres, media);
const outBytes = await zip.generateAsync({ type: "uint8array" });
writeFileSync(resolve(root, ".smoke", "tint-repro.pptx"), outBytes);

// ---- 3) what did we write? ----
const out = await JSZip.loadAsync(outBytes);
console.log("\n=== EXPORTED FILE ===");
for (const name of Object.keys(out.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort()) {
  const xml = await out.file(name).async("text");
  if (!xml.includes("svgBlip")) continue;
  const pics = (xml.match(/<p:pic>[\s\S]*?<\/p:pic>/g) ?? []).filter(p => p.includes("svgBlip"));
  for (const p of pics) {
    const nm = p.match(/name="([^"]*)"/)?.[1];
    if (nm !== target.name) continue;
    console.log(`\n--- ${name} "${nm}" blipFill:`);
    console.log(" ", (p.match(/<p:blipFill>[\s\S]*?<\/p:blipFill>/)?.[0] ?? "").replace(/\s+/g, " "));
    // resolve rels -> media file -> dump svg head
    const relsName = name.replace("slides/", "slides/_rels/") + ".rels";
    const rels = await out.file(relsName).async("text");
    for (const rid of [...p.matchAll(/r:embed="(rId\d+)"/g)].map(m => m[1])) {
      const tgt = rels.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`))?.[1];
      console.log(`  ${rid} -> ${tgt}`);
      if (tgt?.endsWith(".svg")) {
        const svg = await out.file("ppt/" + tgt.replace("../", "")).async("text");
        console.log("  svg head:", svg.slice(0, 400).replace(/\s+/g, " "));
        const fills = [...svg.matchAll(/fill[:=]"?([^;"}]+)/g)].map(m => m[1]).slice(0, 10);
        console.log("  svg fills seen:", fills.join(" | "));
      }
    }
  }
}
console.log("\nDone — .smoke/tint-repro.pptx written (open in PowerPoint to confirm).");
