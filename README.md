# Presentation Editor

A browser-based **PowerPoint (`.pptx`) editor**. It runs entirely client-side
(React + TypeScript + Vite) with a native OOXML/PresentationML engine that reads
and writes **real `.pptx` files** — open a deck made in PowerPoint, edit it, and
save a `.pptx` that opens cleanly back in PowerPoint, LibreOffice, or any
OOXML-compatible suite. No server.

## Run it

```bash
npm install
npm run dev        # http://localhost:5173
npm run build      # static bundle → dist/  (no server needed)
npm run smoke      # .pptx round-trip engine test → "SMOKE OK"
```

Open `http://localhost:5173/embed-test.html` for the iframe-embed demo.

## What it does

Open & save real `.pptx`; edit text (full run/paragraph formatting, theme
fonts), **187 preset shapes** + freeform geometry, images (replace, interactive
crop, SVG recolor), **tables** (PowerPoint-style design gallery, merge/split,
layout ops), **charts** (8 kinds with formatting + per-element text styling),
**groups** (move/resize/rotate/flip as one), gradient/picture/pattern fills, line
arrowheads & dashes, custom color palettes & fonts, self-hosted bundled fonts,
find & replace, rulers, present mode, undo/redo, autosave.

It's embeddable in a host app (open from a URL, save back via presigned PUT or
`postMessage`).

## Documentation

Start at **[CLAUDE.md](CLAUDE.md)** — the documentation index. It links to
[`docs/`](docs/) for architecture, the document model, the OOXML engine,
rendering, state, the UI, the full feature catalog, testing, and an honest
[code-quality assessment](docs/12-code-quality.md). Host-app integration and
deployment are in **[INTEGRATION.md](INTEGRATION.md)**.

## Keyboard

| | |
|---|---|
| F5 / Shift+F5 | Present from start / current slide |
| Esc | Cancel tool · clear selection · leave text/slideshow |
| F2 or double-click | Edit text |
| Ctrl+Z / Ctrl+Y | Undo / redo |
| Ctrl+C/X/V, Ctrl+D | Copy / cut / paste / duplicate |
| Ctrl+G / Ctrl+Shift+G | Group / ungroup |
| Arrows (+Shift) | Nudge 1px (10px) |
| Ctrl+S | Save .pptx |
| PageUp / PageDown | Previous / next slide |
