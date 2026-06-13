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
| `undo()` / `redo()` | void | Step the history. |

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
- **Writes are active-slide-scoped** (see above). Reads are not.
- `setImage` fetches the URL in the browser — the host (R2/S3/CDN) must send
  permissive **CORS** headers, exactly like opening a `?file=` document.
- For **structural** edits the API doesn't cover yet (add slide, insert a new
  shape/chart/table, edit table cells, change chart data), drive the document via
  `pe:load`/`pe:save` round-trips, or open an issue — the surface is designed to
  grow. The internal store ([`src/state/store.ts`](../src/state/store.ts)) has
  the full method set these wrap.

See also: [09 · Integration & embedding](09-integration.md) ·
[INTEGRATION.md](../INTEGRATION.md) (the full host-app guide).
