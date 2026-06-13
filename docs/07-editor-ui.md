# 07 · Editor UI

[← Index](../CLAUDE.md) · Source: [`src/components/`](../src/components/), [`src/App.tsx`](../src/App.tsx)

The UI subscribes to the [store](06-state-store.md), renders the current
presentation, and calls store methods on interaction. It mirrors PowerPoint's
layout: ribbon on top, slide thumbnails on the left, canvas in the middle, a
context settings pane on the right, status bar at the bottom.

## Shell — `App.tsx`

Owns the overall layout and the cross-cutting concerns:

- **Layout grid:** ribbon · (left rail + thumbnails + canvas/notes + right pane) · status bar.
- **Global keyboard:** F5 present, Ctrl+S save, Ctrl+F find, Ctrl+Z/Y undo/redo,
  Ctrl+C/X/V/D clipboard, Ctrl+A select-all, Ctrl+G/Shift+G group, arrows nudge,
  Delete, Page Up/Down.
- **Autosave:** debounced IndexedDB snapshot on every commit.
- **Embed bootstrap:** reads URL config, loads `?file=`, wires the `postMessage`
  bridge, routes Save to upload/host (see [Integration](09-integration.md)).
- **File open/save** and the image file inputs.

## Components

| Component | Role |
|---|---|
| `Ribbon.tsx` | the tabbed toolbar (File/Home/Insert/Design/Transitions/View). Text formatting, shape gallery, chart gallery, theme + font galleries, custom-palette/font dialogs. Home-tab font controls also route to a selected **chart text element**. |
| `EditorCanvas.tsx` | the canvas — selection, drag/move, resize, rotate (single + group), marquee, draw-new-shape, interactive image crop, chart-part selection, rulers, the right-click context menu. |
| `RightPanel.tsx` | the context settings pane — slide/shape/image/table/chart panes, the fill editor (solid/gradient/picture/pattern), the gradient-stops slider, the table design gallery, the accordion-grouped chart settings. |
| `SlidePanel.tsx` | thumbnail rail with drag-reorder, memoized thumbnails, slide context menu, Delete-to-remove. |
| `TextEditOverlay.tsx` | contentEditable overlay for editing shape/table-cell text; parses computed styles back into runs preserving theme color/font refs. |
| `Dropdown.tsx` | viewport-flipping dropdown with a `fixed` mode that escapes clipping ancestors. |
| `ColorPicker.tsx` | theme/standard/custom color grid + the split color button. |
| `CustomizeDialogs.tsx` | "Customize Colors" (12 theme slots) and "Customize Fonts" dialogs. |
| `ChartDataDialog.tsx` | spreadsheet-lite editor for a chart's categories/series. |
| `FindDialog.tsx`, `PatternDialog.tsx`, `NotesBar.tsx`, `PresentMode.tsx`, `StatusBar.tsx`, `icons.tsx` | find & replace, pattern import, speaker notes, slideshow, status bar, hand-drawn SVG icon set. |

## Canvas interaction model

`EditorCanvas` is where the pointer logic lives. A `dragRef` holds a tagged union
of drag kinds — `move`, `resize`, `rotate`, `marquee`, `draw`, `crop`, `croppan`,
`cellsel`, `gresize` (group resize), `grotate` (group rotate). On pointer down it
decides which based on what was hit (a handle, a shape, a chart part, empty
canvas); pointer move drives a live `preview`; pointer up calls `endPreview` so
the gesture is one undo step.

Notable details handled here:

- **Pointer-capture dblclick fix:** `setPointerCapture` retargets the derived
  dblclick to the SVG, so a `lastDownHit` ref records the real target.
- **Group drill-in:** a stationary click on a fully-selected group selects the
  individual member (checked on pointer-up, so dragging the group still works).
- **Chart parts** and **table cell ranges** get their own clickable overlays in
  the selection-chrome layer.
- **Rulers** use `minmax(0,1fr)` grid tracks so the fixed-size ruler SVGs can't
  freeze the canvas size when the layout changes.

## Right pane

Auto-activates the pane matching the selection (slide → shape → image → table →
chart). The chart pane is **accordion-grouped** (Type & options / Chart elements /
Chart area & gridlines / Series / Colors / Text elements / Data & layout), with
open/closed state persisted in localStorage.

## Patterns worth knowing

- **Dropdowns and flyouts** that live inside the scrollable right pane use a
  `fixed`-positioned panel measured from the button, so they escape the pane's
  clipping. (Several "z-index" bugs were really clipping bugs fixed this way.)
- **Text editing** is a separate contentEditable overlay, not inline SVG; on
  commit it maps computed styles back to model runs, preserving symbolic theme
  color/font references.

---

Next: [Feature catalog →](08-features.md)
