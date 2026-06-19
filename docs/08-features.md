# 08 · Feature catalog

[← Index](../CLAUDE.md)

Every user-facing feature, what it does, and where it lives. All visual features
round-trip to `.pptx` unless noted.

## Slides
- Add (with layout), duplicate, delete, drag-reorder, move to start/end. `store` slide ops + `SlidePanel.tsx`.
- Slide size, background (any fill type), apply-to-all. Transitions: fade / push (subset).
- Speaker notes (`NotesBar.tsx`) export as `notesSlides` + `notesMaster`.

## Text
- Font family (system + bundled), size, bold/italic/underline/strike, color, highlight.
- Super/subscript, change case (length-preserving), bullets/numbering, indent, line spacing, vertical align, text columns, clear formatting.
- `Ribbon.tsx` controls → `store.formatRuns`/`formatParagraphs`/`transformText`; editing via `TextEditOverlay.tsx`.

## Shapes
- **187 ECMA-376 preset geometries** in a categorized gallery (Basic/Arrows/Math/Flowchart/Stars/Callouts/…) + recently-used. Evaluated at render time so adjust values round-trip.
- Freeform/custom geometry preserved verbatim on export (`rawGeomXml`).
- Fill: solid / gradient (stops slider) / picture / pattern / none; outline color + **weight + dash styles**; **arrowheads** (begin/end: triangle/stealth/open/diamond/oval × size).
- Move, resize (rotation-aware), rotate, flip H/V, align, distribute, z-order, smart-guide center snapping.

### <a id="grouping"></a>Grouping
- Ctrl+G / Ctrl+Shift+G or right-click. Members share a `groupId`; select/move/resize/**rotate/flip** as one unit; click-again drills into a member. Exports as real `p:grpSp`; imported groups stay grouped.

## Images
- Insert from file; **replace from file or URL** (right-click) keeping the frame.
- **Interactive crop** with PowerPoint-style bracket handles (drag edges; pan the image under the frame), plus numeric crop and rounded-corner/shape frames.
- SVG icons render as vectors and **recolor** ("graphics fill", `svgTint`); export dual-embeds a PNG fallback.

## Tables
- PowerPoint **Table Design**: Light/Medium/Dark × accent style gallery, 6 style-option flags (header/total/banded rows, first/last/banded columns), borders (all/outside/none) + color, cell shading.
- **Layout**: insert/delete rows & columns relative to a cell, distribute, merge/split, cell range selection, cell vertical align + margins. Right-click cell menu + table pane.

## Charts
- 8 kinds (column, bar, line, area, scatter, radar, pie, doughnut) with variants (clustered/stacked/100%, markers, smoothing, radar styles).
- Data editor; series/slice colors; legend position; data labels; error bars; axis titles; hidden axes.
- **Data labels** (`c:dLbls`): to change a label's text, edit its value in the data editor (the label *is* the value). **Position** (`c:dLblPos`) — Outside End / Center / Inside End (Above / Center / Below on line/area/scatter) — via the chart-elements **Data Labels** flyout; mapped per chart type on write so PowerPoint never repair-strips it.
- **Format**: chart-area & plot fill/border, gridline color/visibility, series line width/dash, marker size.
- <a id="charts"></a>**Per-element text styling**: select the title, an axis title, the legend, the axis labels, or the **data labels** (click on the chart or pick in the pane) and format font/size/color/B-I-U from the **Home tab** — exactly PowerPoint's workflow. Data-label styling round-trips as `<c:txPr>` inside `<c:dLbls>`.

## Theming & fonts
- Color themes (Design tab) + **custom palettes** ("Customize Colors", 12 slots, saved).
- Heading/body **theme font pairs** + **custom font pairs** (with optional Google Fonts loading).
- **Self-hosted bundled fonts** (14 families) — see [Build & release](10-build-release.md#fonts).

## Editing aids
- Undo/redo (snapshot history), copy/paste/duplicate, **find & replace**, rulers, zoom (fit/fit-width/numeric), present mode, smart guides, context menus.
- **Autosave** to IndexedDB (the recovery-banner UI was removed for embedding; saving still runs).

## Import/export
- Open & save real `.pptx`. Import **SlideBazaar pattern JSON** (`File → Import Pattern`).
- Host integration: open from URL, save back via presigned PUT or `postMessage` (see [Integration](09-integration.md)).

## Known gaps (deliberate, v1)
Picture *fills* on shapes flatten to gray when no media; slide master/layout
artwork isn't rendered beyond placeholder geometry/text styles; charts use
literal data (no embedded spreadsheet); no real-time co-editing; no PDF export;
no animations beyond fade/push; no accessibility (ARIA) or i18n yet.

---

Next: [Integration & embedding →](09-integration.md)
