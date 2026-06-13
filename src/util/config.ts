/**
 * Host UI configuration — show/hide editor chrome without rebuilding.
 *
 * Resolution order (later wins):
 *   1. defaults (everything on)
 *   2. window.presentationEditorConfig.ui   ← set by the static public/config.js
 *      file the host edits (loaded before the app)
 *   3. URL override  ?ui=fileMenu:0,present:0,tabs.insert:0   ← per-embed
 *
 * Read once at startup. Components import `uiConfig` and gate their rendering;
 * the scripting API exposes the resolved config via getConfig().
 */

export interface UIConfig {
  ribbon: boolean;            // the whole top toolbar
  fileMenu: boolean;          // the File tab + its menu
  save: boolean;              // Save / Download .pptx (toolbar button + menu item)
  open: boolean;              // Open a .pptx
  export: boolean;            // Export as PDF / PNG menu items
  importPattern: boolean;     // Import pattern (JSON)
  newPresentation: boolean;   // Create new
  present: boolean;           // the Present (slideshow) button
  leftRail: boolean;          // left icon rail (find / thumbnails / about)
  slidePanel: boolean;        // slide thumbnail panel
  rightPanel: boolean;        // the format / settings panel
  statusBar: boolean;         // bottom status bar
  notesBar: boolean;          // speaker-notes bar
  docTitle: boolean;          // the editable presentation-title field
  tabs: { home: boolean; insert: boolean; design: boolean; transitions: boolean; view: boolean };
}

type FlagKey = keyof Omit<UIConfig, "tabs">;
type TabKey = keyof UIConfig["tabs"];

const FLAG_KEYS: FlagKey[] = [
  "ribbon", "fileMenu", "save", "open", "export", "importPattern", "newPresentation",
  "present", "leftRail", "slidePanel", "rightPanel", "statusBar", "notesBar", "docTitle",
];
const TAB_KEYS: TabKey[] = ["home", "insert", "design", "transitions", "view"];

export const DEFAULT_UI_CONFIG: UIConfig = {
  ribbon: true, fileMenu: true, save: true, open: true, export: true, importPattern: true,
  newPresentation: true, present: true, leftRail: true, slidePanel: true, rightPanel: true,
  statusBar: true, notesBar: true, docTitle: true,
  tabs: { home: true, insert: true, design: true, transitions: true, view: true },
};

/** Coerce 1/0/true/false (boolean or string) to a boolean, else undefined (= leave default). */
function asBool(v: unknown): boolean | undefined {
  if (v === true || v === false) return v;
  if (v === "1" || v === "true") return true;
  if (v === "0" || v === "false") return false;
  return undefined;
}

/** Pure: merge a raw config object + URL `?ui=` overrides over the defaults. */
export function parseUIConfig(raw: unknown, search = ""): UIConfig {
  const out: UIConfig = { ...DEFAULT_UI_CONFIG, tabs: { ...DEFAULT_UI_CONFIG.tabs } };
  const r = (raw && typeof raw === "object" ? raw : {}) as Record<string, unknown>;

  for (const k of FLAG_KEYS) { const b = asBool(r[k]); if (b !== undefined) out[k] = b; }
  if (r.tabs && typeof r.tabs === "object") {
    const t = r.tabs as Record<string, unknown>;
    for (const k of TAB_KEYS) { const b = asBool(t[k]); if (b !== undefined) out.tabs[k] = b; }
  }

  const ui = new URLSearchParams(search).get("ui");
  if (ui) {
    for (const part of ui.split(",")) {
      const [rawKey, rawVal] = part.split(":");
      const b = asBool(rawVal?.trim());
      const key = rawKey?.trim();
      if (b === undefined || !key) continue;
      if (key.startsWith("tabs.")) {
        const tk = key.slice(5) as TabKey;
        if ((TAB_KEYS as string[]).includes(tk)) out.tabs[tk] = b;
      } else if ((FLAG_KEYS as string[]).includes(key)) {
        out[key as FlagKey] = b;
      }
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Permissions — authoritative gates the EDITOR enforces (not just UI visibility).
// `export` controls whether exportPDF/exportSlidePNG/exportPNGZip are allowed at
// all; when false the API methods reject on every path (UI, bridge, console).
// ---------------------------------------------------------------------------

export interface Permissions {
  export: boolean;
}

const DEFAULT_PERMISSIONS: Permissions = { export: true };

/** Pure: merge a raw permissions object + URL `?perm=` overrides over the defaults. */
export function parsePermissions(raw: unknown, search = ""): Permissions {
  const out: Permissions = { ...DEFAULT_PERMISSIONS };
  const r = (raw && typeof raw === "object" ? raw : {}) as Record<string, unknown>;
  const fileVal = asBool(r.export);
  if (fileVal !== undefined) out.export = fileVal;

  const perm = new URLSearchParams(search).get("perm");
  if (perm) {
    for (const part of perm.split(",")) {
      const [k, v] = part.split(":");
      const b = asBool(v?.trim());
      if (b !== undefined && k?.trim() === "export") out.export = b;
    }
  }
  return out;
}

/** Accept only http(s) URLs (for the optional export-authorization endpoint). */
function httpsUrl(u: unknown): string | undefined {
  if (typeof u !== "string") return undefined;
  try { const abs = new URL(u).href; return /^https?:\/\//i.test(abs) ? abs : undefined; } catch { return undefined; }
}

type RawConfig = { ui?: unknown; permissions?: unknown; exportAuthUrl?: unknown };
const rawConfig: RawConfig =
  (typeof window !== "undefined"
    ? (window as unknown as { presentationEditorConfig?: RawConfig }).presentationEditorConfig
    : undefined) ?? {};
const search = typeof location !== "undefined" ? location.search : "";

/** Resolved UI visibility config, read once at startup. */
export const uiConfig: UIConfig = parseUIConfig(rawConfig.ui, search);

/**
 * Resolved permissions, read once at startup into a module-scoped object the
 * page console cannot reach — so a free-tier user can't flip it from devtools.
 * (A reload re-reads config.js, the host's server-served source of truth.)
 */
export const permissions: Permissions = parsePermissions(rawConfig.permissions, search);

/**
 * Optional per-export server authorization endpoint. When set, the editor POSTs
 * to it (with credentials) before each export and proceeds only on a 2xx. Set it
 * in config.js (server-served) — not the URL — so users can't point it elsewhere.
 */
export const exportAuthUrl: string | undefined = httpsUrl(rawConfig.exportAuthUrl);
