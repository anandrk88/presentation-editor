# 03 · Document model

[← Index](../CLAUDE.md) · Source: [`src/model/types.ts`](../src/model/types.ts), [`src/model/defaults.ts`](../src/model/defaults.ts)

The model is **plain data** — interfaces and discriminated unions, no methods.
It is the single source of truth that the engine, renderer, and store all share.

## Units

All geometry is in **EMU** (English Metric Units): `914400` per inch, `9525` per
CSS pixel at 96 dpi, `12700` per point. Constants live in `defaults.ts`
(`EMU_PER_INCH`, `EMU_PER_PX`, `EMU_PER_PT`). Storing EMU keeps the model
lossless against the file format; the renderer converts to px.

## Colors — `ColorRef`

A color is **either** a literal sRGB hex **or** a reference into the theme's
color scheme, both with optional luminance/alpha modifiers:

```ts
type ColorRef =
  | { kind: "srgb";   hex: string;  lumMod?; lumOff?; alpha? }
  | { kind: "scheme"; slot: SchemeSlot; lumMod?; lumOff?; alpha? };
```

Scheme references (`accent1`, `dk1`, …) are kept **symbolic** so that changing
the theme recolors everything that referenced it. They're flattened to a real
CSS color only at render time by `resolveColor()` (which applies lumMod/lumOff
as HSL adjustments and alpha as rgba).

## Fills — `Fill`

A discriminated union covering every PowerPoint fill type:

```ts
type Fill =
  | { kind: "none" }
  | { kind: "solid";    color: ColorRef }
  | { kind: "gradient"; stops: GradientStop[]; angle: number }
  | { kind: "image";    mediaId: string; tile? }
  | { kind: "pattern";  prst: string; fg: ColorRef; bg: ColorRef };
```

## Shapes — `Shape`

The slide content is `Shape[]`, a union of four kinds sharing `ShapeBase`
(`id, name, x, y, w, h, rot, flipH?, flipV?, groupId?`):

| `kind` | Type | Notable fields |
|---|---|---|
| `"sp"` | `SpShape` | `geom` (preset name), `adj` (adjust values), `custPath`/`rawGeomXml` (freeform), `fill`, `line` (incl. dashes + arrowheads), `text`, `isTextBox` |
| `"pic"` | `PicShape` | `mediaId`, `srcRect` (crop), `geom`/`adj` (frame clip), `svgTint` |
| `"table"` | `TableShape` | `colW`/`rowH`, `cells[][]`, style flags (`firstRow`, `bandRow`, `totalRow`, `firstCol`, `lastCol`, `bandCol`), `styleFamily`/`accent`, `borderMode` |
| `"chart"` | `ChartShape` | `chart` (kind), `categories`, `series[]`, formatting fields, `partStyles` (per-element text), `pointColors`, gridlines |

Text is a `TextBody` (`paragraphs[]`, `anchor`, `wrap`, `insets`, optional
columns); each `Paragraph` has `runs[]`; each `Run` carries
`text/bold/italic/underline/sizePt/font/color` plus optional
`baseline`/`highlight`. Fonts keep OOXML sentinels (`+mj-lt`/`+mn-lt`) symbolic,
mirroring `ColorRef`.

`groupId` is a lightweight grouping mechanism: shapes sharing a `groupId` select,
move, resize, and rotate as a unit and export inside a real `p:grpSp` (see
[Features](08-features.md#grouping)).

## Presentation

```ts
interface Presentation {
  slideWidth: EMU; slideHeight: EMU;   // 16:9 default = 12192000 × 6858000
  slides: SlideModel[];
  theme: ColorTheme;                   // 12 color slots + major/minor font
  title: string;
}
```

A `SlideModel` is `{ id, background?, shapes[], transition?, notes? }`. Media
(images) is **not** in the presentation — it lives in a separate
`Map<string, MediaItem>` keyed by `mediaId`, held by the store, so the same image
is shared and isn't deep-cloned on every undo snapshot.

## `defaults.ts` — themes, resolution, factories

This module is the model's "standard library":

- **Themes:** `OFFICE_THEME`, `GROWTH_THEME`, `THEME_PRESETS`, `THEME_FONT_PAIRS`;
  `MAJOR_FONT`/`MINOR_FONT` sentinels; `resolveColor`, `resolveFill`,
  `resolveFontName` (the symbolic → concrete step).
- **Table styles:** `tableCellStyle(t, r, c)` and `tableBorderColor(t)` compute
  the Light/Medium/Dark × accent look per cell.
- **Factories:** `makeShape`, `makeTextBox`, `makeTable`, `makeChart`,
  `makeSlide`, `tableCell`, `newPresentation`. Every factory seeds a
  **style-carrier run** even for empty text — a hard-won detail: without it,
  typed text inherits the UI font and commits at the wrong size (see the text
  re-edit fix in [Features](08-features.md)).
- **Pickers:** `FONT_LIST` (system + self-hosted bundled families), `FONT_SIZES`.

## Why no classes / behavior

Keeping the model as inert data is what lets the engine and renderer be pure
functions of it, lets undo be a snapshot of it, and lets it serialize to/from
`.pptx` without adapters. Behavior lives in the store (edits) and the
render/engine layers (interpretation).

---

Next: [OOXML engine →](04-ooxml-engine.md)
