# Presentation Editor — Installation & Integration Guide

> **Audience:** a developer or AI agent integrating this editor into a host web
> application (including as a replacement for a server-based slide editor). This
> document is self-contained: everything needed to install, deploy, integrate,
> and verify is here. No other context is required.

---

## 1. What this is

A browser-based **PowerPoint (.pptx) editor** — a presentation-only slide editor
with a native OOXML engine:

- **100% client-side.** React 18 + TypeScript + Vite, compiled to static files.
  Opening, editing, rendering, and exporting `.pptx` all run in the user's
  browser (JSZip + a native PresentationML reader/writer). **There is no
  document server**, no conversion service, no WebSocket backend.
- **Real `.pptx` in/out.** Files exported here open cleanly in Microsoft
  PowerPoint; files authored in PowerPoint open here (text styles, theme
  colors/fonts, 187 preset shapes, freeforms, images incl. SVG icons, crops,
  tables with styles/merges, charts of 8 kinds, groups, speaker notes,
  gradients/patterns/picture fills, transitions subset).
- **Editing features:** ribbon UI (modern-light look), shapes
  gallery, text formatting, find & replace, tables (PowerPoint-style design
  gallery, cell merge, range selection), charts (data editor + formatting),
  image crop with handles, grouping (move/resize/rotate/flip as one), custom
  color palettes & font pairs, autosave with session recovery, present mode,
  rulers, undo/redo.

### Consequences of the no-server architecture (vs a server-based editor)

| | Typical server-based editor | This editor |
|---|---|---|
| Server component | Document server required | None — static files only |
| Concurrent users | Often connection-capped per license | Unlimited (static hosting) |
| Real-time co-editing | Yes | **No** — one editor per document at a time |
| File formats | docx/xlsx/pptx + conversions | **pptx only** |
| Deployment | Docker/server install | Any static host / CDN |

If simultaneous co-editing of one file is required, this editor is not a fit
without adding a sync layer. For open → edit → save-back workflows it is a
drop-in replacement.

---

## 2. Install & deploy

### Prerequisites
- Node.js 18+ and npm.

### Build
```bash
npm install
npm run build        # outputs the entire app to dist/ (static files, ~650 KB JS)
```

### Verify the build (recommended before deploying)
```bash
npm run smoke        # ~260 assertions: full .pptx write→read round-trip engine test
                     # must end with "SMOKE OK"
npm run dev          # dev server at http://localhost:5173 for manual checks
```

### Deploy
Copy `dist/` to any static host (Cloudflare Pages, S3 + CloudFront, nginx,
Vercel, …). The app is a single page; no server-side routing needed.

> **Security requirement:** serve the editor from **its own origin/subdomain**
> (e.g. `https://editor.yourapp.com`), *not* from the host app's origin. The
> editor opens untrusted user files; origin isolation keeps a hypothetical
> exploit away from the host app's cookies/storage. Recommended response
> headers for the editor origin:
> ```
> Content-Security-Policy: default-src 'self'; img-src 'self' data: blob:;
>   style-src 'self' 'unsafe-inline' https://fonts.googleapis.com;
>   font-src https://fonts.gstatic.com data:;
>   connect-src 'self' https://<your-storage-domain>;
> X-Frame-Options: (omit — the host app iframes this) — instead use
> Content-Security-Policy: frame-ancestors https://yourapp.com;
> ```
> (`fonts.googleapis.com`/`gstatic` only needed if users load Google Fonts via
> the Design tab; storage domain needed for `?file=`/`saveUrl` fetches.)

---

## 3. Integration option A — URL parameters (simplest; best for R2/S3)

Embed the editor in an iframe and pass everything in the URL:

```html
<iframe
  src="https://editor.yourapp.com/?file=<ENC_GET_URL>&title=Quarterly%20Deck&saveUrl=<ENC_PUT_URL>"
  style="width:100%;height:100%;border:0"
  allow="clipboard-read; clipboard-write"
></iframe>
```

| Param | Required | Meaning |
|---|---|---|
| `file` | no | URL of the `.pptx` to open on boot. The editor `fetch`es it (CORS must allow the editor origin). Use a short-lived **signed GET URL**. URL-encode it. |
| `title` | no | Document title (header + export filename). |
| `saveUrl` | no | **Presigned PUT URL**. When present, the in-app Save button and `Ctrl/⌘+S` upload the edited `.pptx` to it (`Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation`). When absent (and not in embed mode), Save downloads the file locally. |
| `embed` | no | `1` enables the postMessage bridge (option B). |
| `parentOrigin` | with `embed` | The host app's origin. The bridge refuses all messages from any other origin and posts only to this origin. |

Only `http(s)` URLs are accepted for `file`/`saveUrl` (anything else —
`javascript:`, `ftp:`, etc. — is silently dropped). Relative URLs resolve
against the editor origin.

### Full R2 / S3 flow
1. Host backend mints a signed **GET** URL for the object and a signed **PUT**
   URL for the same key (or a new revision key).
2. Host renders the iframe with `?file=<GET>&saveUrl=<PUT>&title=<name>`.
3. User edits. **Save** PUTs the new bytes to `saveUrl`. The status bar shows
   "Saved to storage"; the document's unsaved-changes flag clears.

#### Backend snippets for minting URLs

**Cloudflare R2 / AWS S3 (Node, @aws-sdk/v3 — R2 is S3-compatible):**
```js
import { S3Client, GetObjectCommand, PutObjectCommand } from "@aws-sdk/client-s3";
import { getSignedUrl } from "@aws-sdk/s3-request-presigner";

const s3 = new S3Client({
  region: "auto", // R2: "auto"; S3: your region
  endpoint: process.env.R2_ENDPOINT, // R2 only: https://<account>.r2.cloudflarestorage.com
  credentials: { accessKeyId: process.env.KEY, secretAccessKey: process.env.SECRET },
});

const getUrl = await getSignedUrl(s3,
  new GetObjectCommand({ Bucket: "decks", Key: key }), { expiresIn: 3600 });
const putUrl = await getSignedUrl(s3,
  new PutObjectCommand({
    Bucket: "decks", Key: key,
    ContentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  }), { expiresIn: 3600 });

const editorUrl = `https://editor.yourapp.com/?file=${encodeURIComponent(getUrl)}` +
  `&saveUrl=${encodeURIComponent(putUrl)}&title=${encodeURIComponent(deckName)}`;
```

**Bucket CORS (required — applies to R2 and S3):**
```json
[
  {
    "AllowedOrigins": ["https://editor.yourapp.com"],
    "AllowedMethods": ["GET", "PUT"],
    "AllowedHeaders": ["Content-Type"],
    "MaxAgeSeconds": 3600
  }
]
```

---

## 4. Integration option B — postMessage bridge (programmatic control)

Use when the host wants the document bytes back itself (store in a DB, upload
through its own API, version it, …).

Embed with the bridge enabled:
```html
<iframe id="editor"
  src="https://editor.yourapp.com/?embed=1&parentOrigin=https%3A%2F%2Fyourapp.com"></iframe>
```

### Messages the HOST sends to the editor
Send with `editor.contentWindow.postMessage(msg, EDITOR_ORIGIN)`.

| `type` | Payload | Effect |
|---|---|---|
| `pe:load` | `{ url: string, title?: string }` **or** `{ data: ArrayBuffer, title?: string }` | Open a presentation (by URL fetch or raw bytes). |
| `pe:save` | — | Export now. Editor replies with `pe:document` (and also PUTs to `saveUrl` if one was set). |

### Events the EDITOR sends to the host
All carry `source: "presentation-editor"` — filter on it. All are posted only
to `parentOrigin`.

| `type` | Payload | When |
|---|---|---|
| `pe:ready` | — | Bridge is listening (iframe booted). |
| `pe:loaded` | `{ title, slideCount }` | A document finished opening (any path: `?file=`, `pe:load`, manual open). |
| `pe:dirty` | `{ dirty: boolean }` | Unsaved-changes flag changed — drive your "unsaved" indicator. |
| `pe:document` | `{ data: ArrayBuffer, fileName }` | Reply to `pe:save` / Save in embed mode. `data` is the complete `.pptx` (transferable). |
| `pe:saved` | `{ via: "upload" \| "message" }` | Save completed (uploaded to `saveUrl`, or delivered via `pe:document`). |
| `pe:error` | `{ message }` | Load/save failure. |

### Complete host-side example
```js
const EDITOR_ORIGIN = "https://editor.yourapp.com";
const editor = document.getElementById("editor");

window.addEventListener("message", async (e) => {
  if (e.origin !== EDITOR_ORIGIN || e.data?.source !== "presentation-editor") return;
  switch (e.data.type) {
    case "pe:ready":
      editor.contentWindow.postMessage(
        { type: "pe:load", url: signedGetUrl, title: "Quarterly Deck" }, EDITOR_ORIGIN);
      break;
    case "pe:dirty":
      setUnsavedIndicator(e.data.dirty);
      break;
    case "pe:document": {
      // raw .pptx bytes — store however you like
      await fetch("/api/decks/123", { method: "PUT", body: new Blob([e.data.data]) });
      break;
    }
    case "pe:error":
      console.error("Editor error:", e.data.message);
      break;
  }
});

// later, e.g. on your own toolbar button:
function saveDeck() {
  editor.contentWindow.postMessage({ type: "pe:save" }, EDITOR_ORIGIN);
}
```

> Note: in dev builds (React StrictMode) some events can fire twice. Make
> handlers idempotent. Production builds fire once.

---

## 5. Migrating from a server-based document editor

If you're replacing a server-based slide editor (one configured via a
`DocEditor`-style object with a document URL, a server callback URL, and event
callbacks), here is the conceptual mapping:

| Server-based DocEditor concept | This editor |
|---|---|
| document URL | `?file=` param or `pe:load { url }` |
| document title | `?title=` param or `pe:load { title }` |
| server callback URL (server POSTs the saved file) | `?saveUrl=` (browser PUTs the file) **or** the `pe:document` event (host receives the bytes) |
| "document ready" event | `pe:ready` (bridge up) + `pe:loaded` (document open) |
| "document state changed" event | `pe:dirty` |
| download / export API | `pe:save` → `pe:document` |
| document type = slides | implicit — this editor is presentations-only |
| document server install / JWT secret | not needed — there's no server |

Key behavioral difference: a server-based editor typically saves
**server-to-server** via its callback URL after the user closes; this editor
saves **from the browser** when the user clicks Save (or on `pe:save`). If you
need autosave-to-storage, listen to `pe:dirty` and send `pe:save` on your own
debounce.

---

## 6. Testing & acceptance checklist

### Bundled live demo (verifies the whole protocol without any host code)
```bash
npm run dev
```
1. Open **`http://localhost:5173/embed-test.html`** — a miniature "host app"
   that iframes the editor with the bridge enabled.
2. Click **Load sample.pptx** → expect log lines `pe:ready` then
   `pe:loaded {"title":"Sample Deck","slideCount":2}`, and the deck visible.
3. Edit something (drag a shape) → expect `pe:dirty {"dirty":true}` and the
   header showing "● unsaved changes".
4. Click **Save (get document back)** → expect
   `pe:document received — Sample Deck.pptx, ~10000+ bytes` and `pe:saved`.
   `window.__lastDoc` holds the ArrayBuffer; first two bytes are `PK` (zip).

### URL-parameter path
5. Open `http://localhost:5173/?file=/sample.pptx&title=Demo` → the deck opens
   automatically with title "Demo".

### Real-storage round trip (with your bucket)
6. Mint signed GET+PUT URLs for a real `.pptx` in R2/S3 (CORS as in §3).
7. Open `?file=<GET>&saveUrl=<PUT>` → deck loads.
8. Edit, press `Ctrl/⌘+S` → status bar shows **"Saved to storage"**.
9. Download the object directly from the bucket and open it in **Microsoft
   PowerPoint** → it must open without repair prompts and show the edit.

### Engine regression suite
10. `npm run smoke` → must print `SMOKE OK` (round-trips shapes, text styles,
    tables incl. merges/styles, charts incl. formatting, images incl. SVG +
    crop, groups, themes, notes, security limits, embed-config parsing).

### Suggested host-app acceptance criteria
- [ ] Editor loads inside the host iframe with the host's real decks from storage.
- [ ] Save persists to storage; re-opening shows the edit.
- [ ] Saved file opens in PowerPoint with no repair dialog.
- [ ] `pe:dirty` drives the host's unsaved indicator.
- [ ] A second deck opened in a second iframe/tab works independently.
- [ ] Oversized/garbage file at `?file=` shows a friendly error (no hang): the
      parser enforces hard limits (150 MB file, 600 MB expansion, 5000 zip
      entries, 1000 slides, depth-capped groups — tunable via `PARSE_LIMITS`
      in `src/ooxml/parse.ts`).

---

## 7. Known limitations (set expectations)

- **pptx only** — no doc/xls, no PDF export (yet).
- **No real-time co-editing** — last writer wins at the storage layer; the host
  should add its own locking/turn-taking if two users may open the same deck.
- **Fonts are referenced, not embedded** — a deck using "Poppins" renders with
  it only where the font is installed/loaded; the editor can load Google Fonts
  for display (Design → Fonts → Customize), and falls back to Segoe UI/Calibri
  otherwise. PowerPoint behaves the same way for non-embedded fonts.
- **Autosave/recovery is per-browser-origin, single slot** — it protects
  against crashes, but in multi-deck embedding rely on host saves (listen to
  `pe:dirty`, call `pe:save`), not on the recovery banner.
- Unsupported embedded objects (OLE, video, SmartArt) import as labeled
  placeholder rectangles with an import warning; slide masters/layout artwork
  beyond placeholder geometry/text styles is not rendered.
- Custom palettes/fonts created in the Design tab are stored per browser
  (localStorage), not inside the file.

---

## 7b. Fonts (self-hosted, bundled with the app)

The editor ships a set of self-hosted fonts so decks render with the right
typefaces regardless of what's installed on the viewer's machine. They're
bundled at build time and lazy-loaded per glyph use (no upfront cost).

**Add / refresh fonts** — drop `.ttf`/`.otf` files in a folder and run:
```bash
npm run fonts -- "D:\path\to\your\fonts"     # default: D:\Only  office fonts
# then rebuild
npm run build
```
The script (`scripts/install-fonts.mjs`):
- copies Regular / Bold / Italic / BoldItalic (weight 400/700 × normal/italic)
  for each **base** family into `public/fonts/`,
- generates `public/fonts/fonts.css` (`@font-face`, `font-display: swap`),
- writes `src/fonts/bundled.ts` (the family list merged into the font picker).

Width variants (e.g. `*_Condensed`) are skipped by default; pass `--all-widths`
to include them. The editor only distinguishes regular/bold and roman/italic,
so other weights (Medium/SemiBold/…) aren't self-hosted.

Currently bundled: Archivo, Exo, Google Sans, Jost, Lato, Metrophobic,
Montserrat, Noto Sans, Playfair Display, Plus Jakarta Sans, Poppins, Roboto,
Sora, Urbanist (~15 MB).

> **Licensing:** most of these are SIL Open Font License (free to self-host).
> **Google Sans is Google-proprietary** — confirm you have the right to
> redistribute it before deploying publicly; to drop it, delete
> `public/fonts/GoogleSans-*.ttf`, its `@font-face` lines, and its entry in
> `src/fonts/bundled.ts` (or just re-run the script from a folder without it).
> Google Sans alone is ~7.6 MB of the bundle.

Fonts are referenced by name in the `.pptx` (PowerPoint uses the viewer's own
installed copy / substitutes) — bundling here only governs in-editor rendering.
Users can also load additional Google Fonts on the fly via Design → Fonts →
Customize.

## 8. Releasing updates (how new versions reach the server)

The editor is a static bundle, so "release" = build + copy `dist/` to the host.
Because the host app integrates only through the **URL + `pe:*` message
protocol** (not through code), the editor updates independently — no host-app
deploy needed for editor releases.

### One-time setup
1. Push this repo to GitHub (it ships with `.github/workflows/deploy.yml`).
2. Enable a deploy target:
   - **Testing (zero config):** repo Settings → Pages → Source: *GitHub Actions*.
     Every push to `main` deploys to `https://<user>.github.io/<repo>/`.
   - **Production:** in `deploy.yml`, enable the `deploy-cloudflare` or
     `deploy-s3` job (flip its `if:` to `true`) and add the listed secrets.

### Per-release flow
```bash
# make changes, then:
npm run smoke          # engine suite must pass
npm run release        # bumps patch version, pushes commit + tag
# → CI: typecheck → smoke → build → deploy. Done.
```
(Or `npm version minor|major` + `git push --follow-tags` for bigger bumps.)

Every push to `main` runs typecheck + the ~260-assertion round-trip suite
before deploying, so a broken engine never ships.

### How running embeds pick up a new version
- Vite emits **content-hashed assets**; `index.html` is tiny and served
  no-cache (the S3 job sets headers; Cloudflare/GitHub Pages handle it).
  **Every new iframe load gets the new version automatically.** Sessions that
  are already open keep running the old version until reloaded — which is the
  safe behavior (never swap code under an open document).
- Each build writes **`/version.json`** (`{ version, commit, builtAt }`), and
  the `pe:ready` event includes `version`. To prompt long-running sessions to
  refresh, poll from the host:

```js
let current;
setInterval(async () => {
  const { version } = await fetch(EDITOR_ORIGIN + "/version.json", { cache: "no-store" }).then(r => r.json());
  if (current && version !== current) showBanner("A new editor version is available — save and reload.");
  current = version;
}, 5 * 60_000);
```

- **Rollback** = re-run the deploy job from the previous tag (or
  `git checkout v0.1.3 && npm run build` + redeploy `dist/`).
- Use two origins for safety: deploy `main` to a **staging** URL your team
  tests against, and promote tags to the **production** URL your app embeds.

## 9. Sharing the code with the host project

Recommended: keep the editor in its **own repository**; the host app consumes
the **deployed origin**, not the source. The `pe:*` protocol is the stable
contract between them, so the two codebases version independently.

| Model | When | How |
|---|---|---|
| **Separate repo + deployed origin** (recommended) | Default | Host app stores only the editor URL (env var like `EDITOR_ORIGIN`). Updates ship by deploying this repo. |
| Git submodule in the host repo | You want the source vendored & pinned | `git submodule add <editor-repo-url> vendor/presentation-editor`; host CI runs its build and serves `dist/` under `/editor/`. |
| Monorepo package | Editor and app share one repo/CI | Move this folder into the monorepo, wire its `build` into the pipeline, serve `dist/` at a dedicated path/subdomain. |

## 10. Repo map (for maintainers)

```
src/model/        document model + defaults/themes (types.ts, defaults.ts)
src/ooxml/        the engine: parse.ts (pptx→model, hardened), write.ts (model→pptx), pattern.ts
src/render/       SVG renderer (slides, 187 preset geometries, tables, charts)
src/components/   ribbon, canvas (selection/drag/crop/groups), panels, dialogs
src/state/        store.ts — single observable store (window.__store in dev)
src/util/         embed.ts (host bridge), autosave.ts, custom.ts (palettes/fonts), svgTint.ts
public/           embed-test.html (host demo), sample.pptx, patterns/
scripts/          smoke-pptx.mjs + smoke-entry.ts (round-trip test suite)
```

`npm run dev` (port 5173) · `npm run build` → `dist/` · `npm run smoke` → engine suite.
