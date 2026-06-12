# Presentation Editor — an ONLYOFFICE-style OOXML slide editor for the browser

A PowerPoint-style presentation editor that runs entirely in the browser and
reads/writes **real `.pptx` files** (ECMA-376 PresentationML). The UI and
architecture follow the open-source ONLYOFFICE Presentation Editor — see
[ARCHITECTURE.md](ARCHITECTURE.md) for the verified mapping to the
`sdkjs` / `web-apps` sources.

## Run it

```bash
npm install
npm run dev        # http://localhost:5173
npm run build      # production bundle in dist/ (static, no server needed)
npm run smoke      # OOXML round-trip validation (writes .smoke/smoke-output.pptx)
```

## What it does

**Editor chrome (ONLYOFFICE classic-light look)**
- Terracotta `#B75B44` header with editable document name and quick-access save/undo/redo
- Ribbon tabs: **File · Home · Insert · Design · Transitions**, with the static
  cluster (copy/paste/undo/redo, Add Slide with layout menu, Start slideshow) on every tab
- Left icon rail + slide thumbnails: click to select, drag to reorder,
  right-click menu (new/duplicate/delete, move to beginning/end, present from here)
- Right settings rail that auto-opens per selection: slide background, shape
  fill/stroke/size/position/rotation/flips, image settings
- Status bar: slideshow, "Slide N of M", theme name, fit-to-slide / fit-to-width, zoom

**Editing**
- Six slide layouts with real placeholders ("Click to add title…")
- 18 OOXML preset shapes drawn by click-drag, plus text boxes and images
- Select / multi-select (shift, marquee), move with smart center guides,
  8-handle resize (shift = aspect), rotation handle with snapping
- In-place rich text editing (double-click / F2): per-run bold, italic,
  underline, strikethrough, font, size, color; alignment, bullets, numbering,
  vertical anchor, line spacing
- Theme color system: six color schemes on the Design tab recolor every
  theme-bound shape — and survive into the saved file
- Slide transitions (fade, push) with speed/direction, applied in the slideshow
  and written to OOXML
- Full-screen presenting (F5), undo/redo, cut/copy/paste/duplicate, arrow-key
  nudge, Ctrl+S to download

**OOXML engine (browser-only, no server)**
- **Save**: complete valid `.pptx` package — content types, relationships,
  theme, slide master + layout, slides, media — opens in PowerPoint,
  LibreOffice and ONLYOFFICE
- **Open**: parses `.pptx` including slide size, themes (scheme colors + fonts),
  solid/gradient-approximated fills, outlines, preset geometry mapping, groups
  (flattened with transforms), pictures, text with run/paragraph properties,
  transitions — and resolves placeholder geometry through the slideLayout →
  slideMaster inheritance chain, the detail that makes real-world decks land
  in the right place

## Known limits (v1)

- Tables, charts, SmartArt import as labeled placeholder boxes; audio/video not supported
- Gradient/picture fills import approximated to solid; custom geometry falls back to rectangles
- Placeholder shapes re-export as plain positioned shapes (they keep look and position, not `ph` re-bindability)
- Text autofit and exact line-break parity with PowerPoint are approximate (browser text layout)
- No collaboration, comments, notes, or animations beyond slide transitions

## Keyboard

| | |
|---|---|
| F5 / Shift+F5 | Present from start / current slide |
| Esc | Cancel tool · clear selection · leave text/slideshow |
| F2 or double-click | Edit text |
| Ctrl+Z / Ctrl+Y | Undo / redo |
| Ctrl+C/X/V, Ctrl+D | Copy / cut / paste / duplicate |
| Arrows (+Shift) | Nudge 1px (10px) |
| Ctrl+S | Download .pptx |
| PageUp / PageDown | Previous / next slide |
