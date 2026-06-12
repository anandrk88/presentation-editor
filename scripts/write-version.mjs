// Emits dist/version.json after a build — host apps poll this to detect new
// releases (see INTEGRATION.md §"Releasing updates").
import { readFileSync, writeFileSync } from "fs";
import { execSync } from "child_process";

const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8"));
let commit = "";
try { commit = execSync("git rev-parse --short HEAD", { stdio: ["ignore", "pipe", "ignore"] }).toString().trim(); } catch { /* not a git checkout */ }

writeFileSync(
  new URL("../dist/version.json", import.meta.url),
  JSON.stringify({ version: pkg.version, commit, builtAt: new Date().toISOString() }, null, 2),
);
console.log(`version.json → v${pkg.version}${commit ? " (" + commit + ")" : ""}`);
