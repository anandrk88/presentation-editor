# 02 · Architecture

[← Index](../CLAUDE.md)

The application is built in **five layers**, each depending only on the ones
below it. Data flows down for rendering and back up through the store for edits.

```
┌─────────────────────────────────────────────────────────────┐
│  UI layer        src/components/   ribbon, canvas, panels     │  React
│                  src/App.tsx       shell, keyboard, embed      │
├─────────────────────────────────────────────────────────────┤
│  State layer     src/state/        EditorStore (observable)    │  imperative
│                                    undo/redo, selection, edits │  + subscribe
├─────────────────────────────────────────────────────────────┤
│  Render layer    src/render/       model → SVG, geometry, charts│  pure
├─────────────────────────────────────────────────────────────┤
│  Engine layer    src/ooxml/        .pptx ⇄ model              │  pure
├─────────────────────────────────────────────────────────────┤
│  Model layer     src/model/        types + defaults/factories  │  data
└─────────────────────────────────────────────────────────────┘
```

The **model** is the single source of truth. It is a plain data structure (no
classes, no behavior). Everything else is a function of it:

- the **engine** converts between the model and `.pptx` bytes,
- the **render** layer converts the model to SVG,
- the **store** holds the current model and the undo history and exposes edit
  operations,
- the **UI** subscribes to the store, renders it, and calls edit operations.

## Data flow

### Opening a file
```
.pptx bytes ──parsePptx()──▶ { pres, media, warnings } ──store.loadPresentation()──▶ store
                                                                                       │
                                                              UI re-renders via subscribe
```

### Editing
```
user gesture ──▶ component ──▶ store.updateShapes()/formatRuns()/… ──▶ commit(nextPres)
                                                                          │
                                       store notifies listeners ──▶ React re-render ──▶ SVG
```

### Saving
```
store.pres + store.media ──buildPptx()──▶ JSZip ──▶ .pptx Blob ──▶ download / PUT / postMessage
```

## The render pipeline

A slide is rendered to SVG by `SlideView.tsx`, which dispatches per shape:

```
SlideSVG(slide)
  └─ for each shape, by kind:
       sp    → SpView      (preset geometry path(s) + fill/stroke + arrows + text)
       pic   → PicView     (image with crop math + frame clip + svg tint)
       table → TableView   (grid + per-cell style fills + borders + cell text)
       chart → ChartView   (axes/series/legend/labels for 8 chart kinds)
```

Geometry for the 187 preset shapes is produced by `presetGeom.ts`, a runtime
evaluator of the ECMA-376 shape-definition tables (adjust values, guide
formulas, path commands) — so a shape renders correctly at *any* size, and
imported adjust values round-trip. See [Rendering](05-rendering.md).

The same `SlideView` renders the **canvas** (full size), the **slide-panel
thumbnails** (small), and **present mode** — one renderer, three consumers.

## The store as the hub

`EditorStore` (`src/state/store.ts`) is a hand-rolled observable:

- holds `state: EditorState` (the presentation + selection + UI flags),
- `subscribe(fn)` / `getState()` feed React's `useSyncExternalStore`,
- every edit goes through `commit(next)` which pushes the previous presentation
  onto an undo stack,
- drags use `preview(next)` / `endPreview(keep)` so a whole gesture is **one**
  undo entry.

The store is the *only* mutable thing. Components never mutate the model; they
call store methods that return new immutable presentations. See
[State & store](06-state-store.md).

## Why it's shaped this way

- **Pure engine + pure renderer** means both are testable in Node without a
  browser. The smoke suite imports the real `parse`/`write` code and runs full
  round-trips — that's why format fidelity is trustworthy. See [Testing](11-testing.md).
- **One model, many views** keeps the canvas, thumbnails, and present mode in
  perfect sync for free.
- **A single store** (instead of Redux/MobX/context-soup) keeps the state
  surface explicit and debuggable (`window.__store` in dev).
- **SVG, not canvas** makes hit-testing, selection, text layout, and crisp
  scaling straightforward, and lets the export reuse the same geometry.

## Boundaries & gotchas worth knowing

- The model stores **theme references symbolically** (`+mj-lt`, scheme slots) and
  resolves them to concrete colors/fonts only at render time. The parser must
  *not* resolve them, or theme changes stop propagating.
- Geometry is in **EMU** (914400 per inch); render divides by 9525 for CSS px.
- The parser is **namespace-agnostic** (matches by local element name) so it
  tolerates the namespace-prefix variations real files use.

---

Next: [Document model →](03-document-model.md)
