# Presentation Editor — Documentation Index

A browser-based **PowerPoint (`.pptx`) editor**. 100% client-side (React + TypeScript +
Vite), with a native OOXML/PresentationML engine that reads and writes real
`.pptx` packages. No server, no document service — it builds to static files.

This file is the **table of contents**. Each topic below lives in its own file
under [`docs/`](docs/); start here and follow the links.

---

## Quick start

```bash
npm install
npm run dev        # dev server at http://localhost:5173
npm run build      # static bundle → dist/  (+ dist/version.json)
npm run smoke      # ~270-assertion .pptx round-trip engine test → must print "SMOKE OK"
npm run fonts -- "C:\path\to\fonts"   # self-host fonts (see docs/10-build-release.md)
```

Open `http://localhost:5173/embed-test.html` for the iframe-embed demo.

---

## Documentation map

| # | Document | What it covers |
|---|----------|----------------|
| 01 | [Overview](docs/01-overview.md) | What it is, the tech stack, the no-server architecture and its consequences, project layout. |
| 02 | [Architecture](docs/02-architecture.md) | The five layers, data flow (model → engine → render → state → UI), the render pipeline, why it's shaped this way. |
| 03 | [Document model](docs/03-document-model.md) | `src/model` — the in-memory shape of a presentation: EMU units, `ColorRef`, the `Shape`/`Fill` discriminated unions, themes, factories. |
| 04 | [OOXML engine](docs/04-ooxml-engine.md) | `src/ooxml` — parsing and writing `.pptx`, the round-trip contract, the format gotchas that cost real debugging, and the untrusted-input hardening. |
| 05 | [Rendering](docs/05-rendering.md) | `src/render` — the SVG renderer, the 187-shape ECMA-376 geometry evaluator, charts, tables, text. |
| 06 | [State & store](docs/06-state-store.md) | `src/state` — the single observable store, undo/redo, selection, the preview/commit pattern, the full method surface. |
| 07 | [Editor UI](docs/07-editor-ui.md) | `src/components` — ribbon, canvas (drag/resize/rotate/crop/group), panels, dialogs, and the interaction model. |
| 08 | [Feature catalog](docs/08-features.md) | Every user-facing feature, what it does, and where it lives in code. |
| 09 | [Integration & embedding](docs/09-integration.md) | Embedding the editor in a host app (URL params + `postMessage` bridge). Links to the full [INTEGRATION.md](INTEGRATION.md). |
| 10 | [Build & release](docs/10-build-release.md) | Versioning, CI, deploy targets, the font-install pipeline, how updates reach the server. |
| 11 | [Testing & verification](docs/11-testing.md) | The smoke round-trip suite and the live-browser verification methodology. |
| 12 | [Code quality assessment](docs/12-code-quality.md) | An honest review: how the code is structured, what's strong, what's tech debt, and a grade. |
| 13 | [Scripting API](docs/13-scripting-api.md) | Programmatic read/write surface: `window.presentationEditor` (+ cross-origin `pe:invoke`) to inspect the active slide / selection / element properties and mutate text, images, fills, geometry. |
| 14 | [Charts — complete reference](docs/14-chart-authoring.md) | Everything about charts: the `ChartShape` model, inserting, reading data back (`ChartInfo`), updating, and every styling dimension (legend, data labels, series colors/lines/markers, axes, title, grouping, fills, per-element fonts) — each tagged with set-via-API / editor-UI / round-trips, plus a capability matrix and the cross-origin envelope. |

---

## Module map (one line each)

```
src/
  main.tsx              app entry (mounts <App/>)
  App.tsx               top-level shell: layout, keyboard, autosave, embed bootstrap, file open/save
  model/
    types.ts            the document model (Shape/Fill/ColorRef unions, TableShape, ChartShape, …)
    defaults.ts         themes, EMU constants, color/font resolution, shape/chart/table factories
  ooxml/
    parse.ts            .pptx (zip) → document model; hardened against untrusted input
    write.ts            document model → .pptx package (opens in PowerPoint)
    pattern.ts          SlideBazaar pattern-JSON importer (raw OOXML fragments + {{placeholders}})
  render/
    SlideView.tsx       slide → SVG (dispatch to shape/pic/table/chart views)
    base.tsx            shared render helpers (paints, fonts, dashes, arrow markers, text)
    presetGeom.ts       ECMA-376 preset geometry evaluator (187 shapes from spec tables)
    geometry.ts         geometry facade + the categorized shape gallery
    GraphicViews.tsx    table + chart renderers (8 chart kinds, chart layout math)
  state/
    store.ts            single observable EditorStore (undo/redo, selection, all edit ops)
    useStore.ts         React binding via useSyncExternalStore
  components/           ribbon, canvas, panels, dialogs (the UI layer); SlideViewer.tsx = touch read-only mobile viewer
  util/
    embed.ts            host-app integration (URL config + postMessage bridge, incl. pe:invoke)
    api.ts              public scripting API (window.presentationEditor): read/mutate/insert slides, elements, text, images, fills; export
    export.tsx          slide → PNG / multi-page PDF / PNG-zip (offscreen render → canvas, lazy jsPDF)
    config.ts           host UI config (window.presentationEditorConfig + ?ui= chrome; ?view=1 read-only mobile viewer)
    autosave.ts         IndexedDB snapshot save/restore
    custom.ts           user palettes + font pairs (localStorage), Google-font loading
    loadImage.ts        image ingest (file / URL), SVG normalize + PNG fallback
    svgTint.ts          SVG recolor for "graphics fill"
  fonts/bundled.ts      generated list of self-hosted font families
scripts/                smoke suite, font installer, preset generator, version stamper
public/                 config.js (host UI config), embed-test.html, sample.pptx, fonts/, patterns/
```

---

## Other top-level docs

- [INTEGRATION.md](INTEGRATION.md) — full host-app integration & deployment guide (the OOXML-editor-embed reference).
- [README.md](README.md) — short landing page.
