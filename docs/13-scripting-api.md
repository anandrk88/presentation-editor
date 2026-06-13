# 13 · Scripting API — read & change the document programmatically

[← Back to the index](../CLAUDE.md)

A stable, host-facing API for **inspecting** what's on screen (active slide,
current selection, every element's properties — text, image, fill, geometry,
table, chart) and **mutating** it from code (set text, swap an image, recolor,
move/resize, select, delete, undo/redo).

This is the surface you build features on top of — a properties panel, a
"replace logo" button, bulk find-and-recolor, an AI "rewrite this slide" action,
analytics on what the user selected, etc.

> Source: [`src/util/api.ts`](../src/util/api.ts). The bridge that exposes it
> cross-origin lives in [`src/util/embed.ts`](../src/util/embed.ts).

---

## Two ways to call it

| | How | When |
|---|---|---|
| **Global object** | `window.presentationEditor` (or `iframe.contentWindow.presentationEditor`) | Same page, or a **same-origin** iframe. Synchronous, fully typed. |
| **postMessage** | `pe:invoke { requestId, method, args }` → `pe:result { requestId, ok, value }` | **Cross-origin** iframe. Async; every method below is callable by name. |

Both call the *same* methods and return the *same* shapes. Pick the global when
you can (it's simpler); use postMessage when the editor is on another origin.

### Global (same-origin)

```js
const api = window.presentationEditor;          // editor in the same page
// or, from a host page embedding a same-origin iframe:
const api = document.getElementById("editor").contentWindow.presentationEditor;

const sel = api.getSelection();
if (sel.elements[0]?.kind === "shape") {
  api.setText(sel.ids[0], "New title");
}
```

### postMessage (cross-origin)

```js
const editor = document.getElementById("editor").contentWindow;
const EDITOR_ORIGIN = "https://editor.example.com";

function call(method, ...args) {
  const requestId = crypto.randomUUID();
  return new Promise((resolve, reject) => {
    function onMsg(e) {
      if (e.origin !== EDITOR_ORIGIN) return;
      const d = e.data;
      if (d?.source !== "presentation-editor" || d.type !== "pe:result" || d.requestId !== requestId) return;
      window.removeEventListener("message", onMsg);
      d.ok ? resolve(d.value) : reject(new Error(d.error));
    }
    window.addEventListener("message", onMsg);
    editor.postMessage({ type: "pe:invoke", requestId, method, args }, EDITOR_ORIGIN);
  });
}

const selection = await call("getSelection");
await call("setText", selection.ids[0], "New title");
```

> The iframe must be embedded with `&embed=1&parentOrigin=<your origin>` for the
> bridge to accept commands. See [09 · Integration](09-integration.md).

---

## Units

Geometry is reported in **EMU** — the native OOXML unit — with a parallel `px`
block so you don't have to convert:

```
914400 EMU = 1 inch     9525 EMU = 1 CSS px (@96dpi)     12700 EMU = 1 pt
```

`api.EMU_PER_PX`, `api.EMU_PER_INCH`, `api.EMU_PER_PT` are exposed for your own
conversions. Every element carries both: `{ x, y, w, h }` in EMU and
`px: { x, y, w, h }` in CSS px.

---

## Reading

All read methods are synchronous and return plain JSON (safe to `postMessage`).

### `getDocument()` → `DocumentInfo`

```ts
{
  title: string,
  slideCount: number,
  slideWidth: number, slideHeight: number,   // EMU
  activeSlideIndex: number,
  dirty: boolean, canUndo: boolean, canRedo: boolean,
  palette: { dk1, lt1, accent1, …: string }  // theme colors as #hex
}
```

### `getSlides()` → `SlideInfo[]`

A light summary of every slide (no element detail): `{ index, id, title?,
elementCount, background? }`. `title` is the first non-empty line of text on the
slide — a usable label for a slide list.

### `getActiveSlide()` → `SlideInfo` &nbsp;·&nbsp; `getSlide(index)` → `SlideInfo | null`

Like `getSlides()` but **includes `elements: ElementInfo[]`** — the full
properties of every shape on that slide. `getActiveSlide()` is the one slide the
user is looking at; `getSlide(i)` reads any slide by index.

### `getSelection()` → `SelectionInfo`

```ts
{ slideIndex: number, ids: string[], elements: ElementInfo[] }
```

What's selected right now. `ids` is empty when nothing is selected.

### `getElement(id)` → `ElementInfo | null` &nbsp;·&nbsp; `getElements()` → `ElementInfo[]`

One element by id, or all elements on the active slide. Returns `null` if the id
isn't on the active slide.

### `ElementInfo`

The stable projection of one element. Type-specific blocks appear only for the
matching `kind`:

```ts
{
  id: string,
  name: string,
  kind: "shape" | "image" | "table" | "chart",
  groupId?: string,                 // present when the element is in a group

  // geometry (EMU) + the same rectangle in CSS px
  x, y, w, h: number,
  rot: number, flipH: boolean, flipV: boolean,
  px: { x, y, w, h: number },

  // — kind: "shape" —
  text?: string,                    // plain text, paragraphs joined by "\n"
  paragraphs?: [{
    text: string, align: "l"|"ctr"|"r"|"just", level: number,
    bullet: "none"|"char"|"num",
    runs: [{ text, bold, italic, underline, strike, caps?, sizePt, font, color }]
  }],
  geom?: string,                    // preset name: "rect", "ellipse", "hexagon"…
  fill?: FillInfo,                  // see below
  line?: { color: string, widthPt: number, dash?: string },

  // — kind: "image" —
  image?: { mediaId, mime?, dataUrl?, crop?: { l, t, r, b } },

  // — kind: "table" —
  table?: { rows: number, cols: number, cells: string[][] },   // [row][col] plain text

  // — kind: "chart" —
  chart?: { chartType, title?, legend, categories: string[], series: [{ name, values: number[] }] }
}
```

`FillInfo` is one of:

```ts
{ kind: "none" }
{ kind: "solid", color: "#RRGGBB" }
{ kind: "gradient", angle, stops: [{ pos, color }] }
{ kind: "image", mediaId, tile }
{ kind: "pattern", preset, fg, bg }
```

Colors are **already resolved** to CSS strings — theme references like
`accent1` and font refs like `+mn-lt` are substituted, so what you read is what
renders.

---

## Writing

Mutations go through the editor's history (each is **one undo step** and marks
the document **dirty**, which fires the `pe:dirty` event so your autosave runs —
see [09 · Integration](09-integration.md)).

> **Active-slide scope.** Writes target the **active slide** (mirrors the
> editor's own model). To change an element on another slide, call
> `selectSlide(i)` first. Reads (`getSlide`) work on any slide.

### Selection, element properties & history

| Method | Returns | Effect |
|---|---|---|
| `selectSlide(index)` | void | Make slide `index` active (clears element selection). |
| `selectElement(id \| id[])` | void | Select element(s) on the active slide (unknown ids are ignored). |
| `clearSelection()` | void | Deselect everything. |
| `setText(id, text)` | `boolean` | Replace a shape's text. `"\n"` splits paragraphs; the first run's font/size/color is reused. `false` if `id` isn't a text-capable shape. |
| `setElementProperties(id, props)` | `boolean` | Move/resize/rotate/flip. `props`: any of `{ x, y, w, h, rot, flipH, flipV }` (EMU). Omitted fields are left as-is; `w`/`h` clamp to ≥1. |
| `setFillColor(id, hex)` | `boolean` | Solid-fill a shape. `hex` accepts `"#RRGGBB"`, `"RRGGBB"`, or `"RGB"`. `false` if not a shape or the hex is invalid. |
| `setImage(id, urlOrDataUrl)` | `Promise<boolean>` | Replace a picture's image from an http(s) URL or a `data:` URL (CORS must allow it). Keeps the frame; resets crop/tint. `false` if `id` isn't an image. |
| `deleteElement(id)` | `boolean` | Remove an element from the active slide. |
| `reorderElement(id, dir)` | `boolean` | Z-order: `dir` is `"front"`, `"back"`, `"forward"`, or `"backward"`. |
| `undo()` / `redo()` | void | Step the history. |

### Slides & document

| Method | Returns | Effect |
|---|---|---|
| `addSlide(opts?)` | `number` | Insert a slide and return its index. `opts`: `{ layout?, index? }` — `layout` is `"title"`, `"titleContent"`, `"sectionHeader"`, `"twoContent"`, `"titleOnly"`, or `"blank"` (default `"titleContent"`); `index` is where to insert (default: after the active slide). The new slide becomes active. |
| `duplicateSlide(index?)` | `number` | Copy a slide (default: active) and return the new index. |
| `deleteSlide(index?)` | void | Remove a slide (default: active). Never drops below one slide. |
| `moveSlide(from, to)` | void | Reorder slides. |
| `setDocumentTitle(title)` | void | Rename the presentation (used as the export file name). |
| `applyTheme(palette)` | void | Recolor theme scheme slots — `palette` maps any of `dk1, lt1, dk2, lt2, accent1…accent6, hlink, folHlink` to a hex. Theme-referenced colors across the deck update at once. |
| `setSlideBackgroundColor(hex)` | `boolean` | Solid background on the active slide. |

### Insert new elements

Each adds to the **active slide**, selects it, and returns the new element id.
`opts` carries an optional box `{ x, y, w, h }` in EMU; omit it for a centered default.

| Method | Returns | Effect |
|---|---|---|
| `insertText(opts?)` | `string` | New text box. `opts.text` seeds the content. |
| `insertShape(geom, opts?)` | `string` | New preset shape (`geom`: `"rect"`, `"ellipse"`, `"hexagon"`, `"star5"`, …). `opts.fillColor` sets a solid fill. |
| `insertImage(urlOrDataUrl, opts?)` | `Promise<string>` | New picture from a URL / `data:` URL (CORS applies). Sized to its natural aspect if no box is given. |
| `insertChart(chartType, opts?)` | `string` | New chart (`chartType`: `"column"`, `"bar"`, `"line"`, `"pie"`, `"doughnut"`, `"area"`, `"scatter"`, `"radar"`). `opts.categories` / `opts.series: [{ name?, values }]` seed the data. |
| `insertTable(rows, cols, opts?)` | `string` | New `rows × cols` table. |

### Element content & style

| Method | Returns | Effect |
|---|---|---|
| `setTableCell(id, row, col, text)` | `boolean` | Set a cell's text (0-based; `"\n"` splits lines). `false` if out of range or not a table. |
| `setChartData(id, data)` | `boolean` | Replace a chart's `{ categories?, series? }` (`series: [{ name?, values }]`). |
| `setParagraphStyle(id, style)` | `boolean` | Apply to all paragraphs: `{ align?, bullet?, level? }`. |
| `setTextStyle(id, style)` | `boolean` | Apply to all runs: `{ bold?, italic?, underline?, strike?, sizePt?, font?, color? }` (`color` is a hex). |

---

## Export — PDF & PNG

Render slides to images entirely in the browser. Each returns a `Blob` (the
renderer and jsPDF load on demand, so they never weigh down the main bundle).

| Method | Returns | Output |
|---|---|---|
| `exportSlidePNG(index?, opts?)` | `Promise<Blob>` | One slide as a PNG (default: the active slide). |
| `exportPDF(opts?)` | `Promise<Blob>` | The whole deck as a multi-page PDF, one slide per page. |
| `exportPNGZip(opts?)` | `Promise<Blob>` | Every slide as a PNG, bundled into one `.zip`. |

`opts.scale` (default `2`) is device-px per slide-px — raise it for print-grade
output, lower it for smaller files. (A 13.3″ 16:9 slide is 1280×720 slide-px, so
scale 2 → a 2560×1440 PNG.)

```js
// same-origin: download a PDF
const blob = await api.exportPDF();
const url = URL.createObjectURL(blob);
Object.assign(document.createElement("a"), { href: url, download: "deck.pdf" }).click();

// cross-origin: the Blob comes back through pe:result
const pdf = await call("exportPDF");          // a Blob
```

The editor's own **File → Export as PDF / PNG** menu items call these.

> **Font note.** Text is rasterized with whatever font the browser has for each
> family. System/common families render exactly; an exotic bundled family the
> export can't see falls back — the same substitution PowerPoint does for a
> missing font. Geometry, fills, images, charts and tables are pixel-exact.

---

## Events (global only)

`api.on(event, handler)` returns an **unsubscribe** function. Handy for keeping a
host-side panel in sync.

| Event | Payload | Fires when |
|---|---|---|
| `"selection"` | `SelectionInfo` | The selection (or active slide) changes. |
| `"slide"` | `SlideInfo` (active) | The active slide changes. |
| `"dirty"` | `boolean` | The unsaved-changes flag flips. |
| `"change"` | `undefined` | Any document edit is committed. |

```js
const off = api.on("selection", sel => {
  panel.render(sel.elements[0]);     // update your properties panel
});
// later: off();
```

Cross-origin hosts get the same signals pushed automatically as
`pe:selection { selection }` and `pe:slide { slide }` messages (and `pe:dirty`,
already part of the bridge).

---

## Recipes

**Show the selected element's text in your own panel**

```js
api.on("selection", sel => {
  const el = sel.elements[0];
  myPanel.textContent = el?.text ?? "(no text selected)";
});
```

**Replace the logo on every slide**

```js
const total = api.getDocument().slideCount;
for (let i = 0; i < total; i++) {
  api.selectSlide(i);
  for (const el of api.getActiveSlide().elements) {
    if (el.kind === "image" && el.name.toLowerCase().includes("logo")) {
      await api.setImage(el.id, "https://cdn.example.com/new-logo.png");
    }
  }
}
```

**Recolor every red shape to the theme accent**

```js
const accent = api.getDocument().palette.accent1;
for (const el of api.getElements()) {
  if (el.fill?.kind === "solid" && el.fill.color.toUpperCase() === "#FF0000") {
    api.setFillColor(el.id, accent);
  }
}
```

**Read everything on the current slide (cross-origin)**

```js
const slide = await call("getActiveSlide");
console.log(slide.elements.map(e => `${e.kind}: ${e.text ?? e.name}`));
```

---

## Notes & limits

- **Ids are stable within a session** but are regenerated when a deck is loaded
  or duplicated — don't persist them across documents.
- **Writes (element & insert ops) are active-slide-scoped** (see above). Reads,
  and slide/document ops (`addSlide`, `moveSlide`, `applyTheme`, …), are not.
- `setImage` / `insertImage` fetch the URL in the browser — the host (R2/S3/CDN)
  must send permissive **CORS** headers, exactly like opening a `?file=` document.
- `setTextStyle` / `setParagraphStyle` apply to **every** run / paragraph of the
  shape. For per-run edits, drive the deck via `pe:load`/`pe:save` round-trips.
- Not yet wrapped: cell merge/insert-row-col, group/ungroup, per-slice chart
  colors, transitions. The internal store
  ([`src/state/store.ts`](../src/state/store.ts)) has the full method set these
  wrap — the surface is designed to grow.

See also: [09 · Integration & embedding](09-integration.md) ·
[INTEGRATION.md](../INTEGRATION.md) (the full host-app guide).
