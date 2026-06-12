# Architecture — and how it maps to the real ONLYOFFICE

This editor is a clean-room TypeScript clone of the ONLYOFFICE **Presentation Editor**
(PowerPoint-style, OOXML/PresentationML). Before building, the two open-source
ONLYOFFICE repos were studied directly:

- [`ONLYOFFICE/sdkjs`](https://github.com/ONLYOFFICE/sdkjs) — the editor engine (model, canvas renderer, OOXML semantics)
- [`ONLYOFFICE/web-apps`](https://github.com/ONLYOFFICE/web-apps) — the UI shell (`apps/presentationeditor`)

A sparse clone of `web-apps` used during research is kept in `_research/` (safe to delete).

## What we copied from ONLYOFFICE's design

| ONLYOFFICE (verified in source) | This clone |
|---|---|
| UI never touches the model; everything goes through the `Asc.asc_docs_api` facade + `asc_on*` events (`slide/api.js`) | React components call `EditorStore` methods; views subscribe via `useSyncExternalStore` ([store.ts](src/state/store.ts)) |
| Model classes mirror OOXML 1:1 — `CPresentation`, `Slide.cSld.spTree`, `CShape{nvSpPr, spPr, txBody}`, theme = clrScheme+fontScheme (`slide/Editor/Format/*`, `common/Drawings/Format/*`) | [types.ts](src/model/types.ts): `Presentation → SlideModel → Shape(sp/pic)` with EMU geometry, `ColorRef` scheme references, `TextBody → Paragraph → Run` |
| **Placeholder inheritance**: shape without `xfrm` resolves geometry by walking layout → master via `getMatchingShape(type, idx)` with normalization `ctrTitle→title`, `null type→body`, title matches by type (`Slide.js:390`, `Shape.js:1694`) | [parse.ts](src/ooxml/parse.ts) `findPh()` implements the same normalization and slide→layout→master walk for imported decks |
| Model coordinates in physical units (mm), zoom is render-time scale only — no re-layout on zoom | Model in EMU; rendering maps EMU→px once (`/9525`); zoom is an SVG scale factor |
| Static slide rendering separated from an interaction **overlay** (selection frames, handles, ghost "track objects" committed on mouse-up — `common/Overlay.js`, `TrackObjects/*`) | [SlideSVG](src/render/SlideView.tsx) renders the document; [EditorCanvas](src/components/EditorCanvas.tsx) layers hit-rects/handles/marquee/guides above it; drags mutate via `preview()` and only enter undo history on pointer-up (`endPreview`) |
| Preset geometries = ECMA-376 `presetShapeDefinitions.xml` formulas evaluated at runtime (`CreateGeometry.js`, 187 presets) | [geometry.ts](src/render/geometry.ts) evaluates the same spec formulas for an 18-preset subset, emitting SVG paths |
| Undo = history points of reversible change objects (`word/Editor/History.js`) | Snapshot history of the immutable `Presentation` value — simpler, same UX (Ctrl+Z/Y, 200 steps) |
| Text engine: `txBody.content` is a mini rich-text document of paragraphs/runs | Same shape: `TextBody.paragraphs[].runs[]`; in-place editing via a zoomed `contentEditable` overlay parsed back to runs ([TextEditOverlay.tsx](src/components/TextEditOverlay.tsx)) |
| Ribbon: File/Home/Insert/Design/Transitions…, static cluster (copy/paste/undo/redo + **Add Slide** + **Start slideshow**) on every tab; header `#B75B44`; 28/32/66px rows; 40px rails; 260px right panel; 25px statusbar with "Slide N of M" + fit/zoom (`Toolbar.template`, `colors-table.less`, `variables.less`) | [Ribbon.tsx](src/components/Ribbon.tsx) + [styles.css](src/styles.css) reproduce the layout and the verified design tokens; right panel auto-switches per selection like `asc_onFocusObject` ([RightPanel.tsx](src/components/RightPanel.tsx)) |

## What we deliberately do differently

| ONLYOFFICE | This clone | Why |
|---|---|---|
| `.pptx` is converted server-side by the C++ `x2t` into a binary (`PPTY`) the browser consumes; saving reverses it | **Direct `.pptx` read/write in the browser** (JSZip + DOMParser / string serialization, [parse.ts](src/ooxml/parse.ts), [write.ts](src/ooxml/write.ts)) | No server needed; sdkjs itself is moving this way (`OpenDocumentFromZipNoInit`) |
| Canvas 2D rendering with own font shaper (HarfBuzz WASM) | SVG rendering (+ HTML `foreignObject` for wrapped text) | Crisp at any zoom, free hit-testing/selection, one renderer shared by canvas/thumbnails/slideshow — right trade-off at clone scale |
| Lock-based realtime co-editing over binary change streams | Skipped | Out of scope; snapshot history would translate to CRDTs if ever needed |
| Full Word text layout (bidi, shaping, math, autofit) | Browser's text layout | The browser already does this well for a demo editor |

## Source map

```
src/
  model/      types.ts (OOXML-mirroring document model, EMU units)
              defaults.ts (Office theme, layouts/placeholders, color resolution)
  ooxml/      write.ts (model -> complete .pptx package)
              parse.ts (.pptx -> model; placeholder inheritance, groups, themes)
  render/     geometry.ts (ECMA-376 preset shape paths)
              SlideView.tsx (shared SVG slide renderer)
  state/      store.ts (facade/store: selection, history, slide & shape ops)
  components/ Ribbon, SlidePanel, EditorCanvas, TextEditOverlay,
              RightPanel, StatusBar, PresentMode, ColorPicker, Dropdown, icons
scripts/      smoke-pptx.mjs + smoke-entry.ts (round-trip package validation)
```

## OOXML package this editor writes

```
[Content_Types].xml          _rels/.rels              docProps/{core,app}.xml
ppt/presentation.xml (+rels) ppt/{presProps,viewProps,tableStyles}.xml
ppt/theme/theme1.xml         ppt/slideMasters/slideMaster1.xml (+rels, clrMap, txStyles)
ppt/slideLayouts/slideLayout1.xml (+rels)
ppt/slides/slideN.xml (+rels: layout, images)    ppt/media/imageN.*
```

Slides carry `p:sp` (xfrm with rot/flips, prstGeom, solid/no fill, ln, full txBody
with pPr/buChar/buAutoNum and rPr runs), `p:pic` (blipFill → media), `p:bg`, and
`p:transition` (fade/push). Colors are written as `a:schemeClr` (+lumMod/lumOff)
when theme-bound, `a:srgbClr` otherwise — so re-theming works in PowerPoint too.
