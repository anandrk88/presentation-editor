/**
 * OOXML round-trip smoke test (Node):
 *  1. builds a presentation through the real writer (TS sources via vite build output? no — esbuild on the fly)
 *  2. validates package structure + XML well-formedness of every part
 *  3. re-opens the file through the real parser and checks content survived
 */
import { build } from "vite";
import { fileURLToPath, pathToFileURL } from "url";
import { dirname, resolve } from "path";
import { writeFileSync, mkdirSync, readFileSync, existsSync } from "fs";

const root = dirname(dirname(fileURLToPath(import.meta.url)));

// Bundle the engine (writer+parser+model) to a single ESM file we can import.
mkdirSync(resolve(root, ".smoke"), { recursive: true });
await build({
  configFile: false,
  root,
  logLevel: "error",
  build: {
    lib: { entry: resolve(root, "scripts/smoke-entry.ts"), formats: ["es"], fileName: "engine" },
    outDir: resolve(root, ".smoke"),
    emptyOutDir: true,
    rollupOptions: { external: [] },
    minify: false,
  },
});

const { runSmoke } = await import(pathToFileURL(resolve(root, ".smoke", "engine.js")).href + `?t=${Date.now()}`);
const patternPath = resolve(root, "public", "patterns", "growth_report_cover_slide.json");
const patternJson = existsSync(patternPath) ? readFileSync(patternPath, "utf8") : undefined;
const { zipBytes, report } = await runSmoke(patternJson);
writeFileSync(resolve(root, ".smoke", "smoke-output.pptx"), zipBytes);

let failed = 0;
for (const line of report) {
  console.log(line.startsWith("FAIL") ? `\x1b[31m${line}\x1b[0m` : line);
  if (line.startsWith("FAIL")) failed++;
}
console.log(failed === 0 ? "\x1b[32mSMOKE OK\x1b[0m — .smoke/smoke-output.pptx written" : `\x1b[31m${failed} failure(s)\x1b[0m`);
process.exit(failed === 0 ? 0 : 1);
