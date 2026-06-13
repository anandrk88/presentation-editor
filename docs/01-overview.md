# 01 · Overview

[← Index](../CLAUDE.md)

## What it is

A **presentation editor that runs entirely in the browser** and reads/writes
real **PowerPoint `.pptx` files** (the ECMA-376 PresentationML format). You can
open a deck made in PowerPoint, edit it — text, shapes, tables, charts, images,
groups — and save a `.pptx` that opens cleanly back in PowerPoint, LibreOffice,
or any OOXML-compatible suite.

It is presentation-only (no documents, no spreadsheets).

## Tech stack

| Concern | Choice |
|---|---|
| Language | TypeScript (strict) |
| UI | React 18 |
| Build | Vite 5 |
| Rendering | SVG (no canvas, no WebGL) |
| `.pptx` I/O | JSZip + a hand-written PresentationML reader/writer |
| State | a single observable store + `useSyncExternalStore` (no Redux/MobX) |
| Persistence | IndexedDB (autosave), localStorage (user palettes/fonts/UI prefs) |
| Tests | a Node round-trip "smoke" suite (~270 assertions) |

There is **no backend**. `npm run build` produces a folder of static files.

## The no-server architecture (and what it means)

Everything — opening, parsing, rendering, editing, exporting — happens in the
user's browser as JavaScript. Consequences:

- **Unlimited concurrent users.** Every user runs their own copy; the host only
  serves static files. There is no per-connection cap.
- **Trivial deployment.** Any static host or CDN (Cloudflare Pages, S3 +
  CloudFront, nginx, GitHub Pages).
- **No real-time co-editing.** Two people cannot edit the *same* document
  together (there is no sync layer). One editor per document; exchange happens by
  saving and sharing the `.pptx`.
- **Files are the contract.** The app is embedded in a host application, handed a
  `.pptx` URL, and gives the edited file back (see [Integration](09-integration.md)).

## Project layout

```
.
├── CLAUDE.md            ← documentation index (start here)
├── INTEGRATION.md       ← host-app embedding & deployment guide
├── README.md
├── docs/                ← these documents
├── index.html           ← Vite entry; links self-hosted fonts
├── vite.config.ts       ← base "./", injects __APP_VERSION__
├── package.json         ← scripts: dev / build / smoke / fonts / release
├── tsconfig.json
├── .github/workflows/   ← CI: typecheck → smoke → build → deploy
├── public/              ← static assets shipped as-is (fonts, sample.pptx, embed-test.html, patterns)
├── scripts/             ← build/test tooling (Node)
└── src/                 ← application source (see the module map in CLAUDE.md)
```

## Size at a glance

~11,000 lines of source. The largest modules are the UI panel (`RightPanel.tsx`),
the OOXML parser (`parse.ts`), and the canvas (`EditorCanvas.tsx`) — see
[Code quality](12-code-quality.md) for what that implies.

---

Next: [Architecture →](02-architecture.md)
