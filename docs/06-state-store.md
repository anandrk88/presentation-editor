# 06 · State & store

[← Index](../CLAUDE.md) · Source: [`src/state/store.ts`](../src/state/store.ts), [`src/state/useStore.ts`](../src/state/useStore.ts)

All application state lives in a **single observable store**, `EditorStore`. There
is no Redux, MobX, or React context tree — just one class instance (`store`) and
React's `useSyncExternalStore`.

## EditorState

```ts
interface EditorState {
  pres: Presentation;            // the document
  selection: { slideIndex, shapeIds[] };
  editingShapeId, editingCell, tableSel, chartEditId, chartPartSel, croppingId;
  zoom, effectiveZoom, presenting, presentIndex, showRuler, pendingShape;
  dirty, canUndo, canRedo, statusMessage;
}
```

`pres` is the document; everything else is editor/UI state. Media is held
separately in `store.media` (a `Map`) so images aren't deep-cloned on undo.

## React binding

`useStore.ts` is tiny:
```ts
export const useEditorState = () =>
  useSyncExternalStore(store.subscribe, store.getState);
```
Components read `state` and call `store.*` methods. In dev, `window.__store` is
exposed for inspection and scripted testing.

## History: commit / preview / endPreview

The undo model is **snapshot-based**:

- `commit(next)` pushes the *previous* `pres` onto the undo stack and sets the
  new one. Simple, correct, and what most edits call.
- `preview(next)` updates the rendered `pres` **without** touching history — used
  during a drag/resize/rotate so the canvas updates live.
- `endPreview(keep)` finalizes: on the first preview it remembered the pre-drag
  base; on release it pushes that base once. **A whole gesture = one undo entry.**

This is the key pattern that makes interactive editing feel right while keeping
undo coarse-grained. `updateShapes(ids, fn, { historic:false })` is the
preview-path variant used by drag handlers.

## The method surface

The store exposes a flat, explicit API (one method per operation), grouped:

- **Slides:** `selectSlide`, `addSlide`, `duplicateSlide`, `deleteSlide`, `moveSlide`, `setSlideSize`, `setTransition`/`applyTransitionToAll`, `updateSlideBg`/`applyBgToAll`, `addImportedSlide`, `loadPresentation`, `newPresentation`.
- **Shapes:** `selectShapes`, `addShape`, `updateShapes`, `deleteSelectedShapes`, `reorderShape`, `alignSelected`, `distributeSelected`, `copySelection`/`pasteClipboard`, `replacePicMedia`.
- **Groups:** `expandToGroups`, `groupSelection`, `ungroupSelection`, `selectionHasGroup`, `rotateSelected`, `flipSelected`.
- **Text:** `formatRuns`, `formatParagraphs`, `transformText` (length-preserving case changes across run boundaries).
- **Tables:** `modifyTable`, insert/delete row+col, distribute, `mergeTableRight/Down/Cells`, `splitTableCell`, `applyTableStyle`, cell/range fill + anchor + insets.
- **Charts:** `setChartPart`, `formatChartPart`, `resetChartPart`, `chartPartStyle`.
- **Find/replace:** `findMatches`, `replaceAll`.
- **Misc:** `setStatus`, `setState`.

Every mutating method builds a **new immutable** presentation (via the private
`withSlide` helper) and routes through `commit`/`preview` — components never
mutate the model directly.

## Why a single store

- The state surface is **explicit and greppable** — you can read every operation
  in one file.
- Undo/redo is trivially correct because the whole document is immutable and
  snapshotted.
- `useSyncExternalStore` gives tearing-free reads with zero boilerplate.
- It's debuggable: `window.__store.state` in the console is the entire app state.

The trade-off (snapshot undo deep-clones the presentation per commit) is
discussed in [Code quality](12-code-quality.md).

---

Next: [Editor UI →](07-editor-ui.md)
