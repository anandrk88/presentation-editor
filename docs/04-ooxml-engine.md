# 04 · OOXML engine

[← Index](../CLAUDE.md) · Source: [`src/ooxml/parse.ts`](../src/ooxml/parse.ts), [`src/ooxml/write.ts`](../src/ooxml/write.ts), [`src/ooxml/pattern.ts`](../src/ooxml/pattern.ts)

This is the heart of the application: the pure functions that turn `.pptx` bytes
into the [document model](03-document-model.md) and back. Both directions are
verified by the [round-trip smoke suite](11-testing.md).

A `.pptx` is a **ZIP** of XML parts plus media. The engine uses JSZip to read/write
the archive and parses/serializes the XML.

## Parsing — `parsePptx(data, fileName)`

Returns `{ pres, media, warnings }`. The flow mirrors how PowerPoint loads a file:

```
/_rels/.rels ─▶ ppt/presentation.xml ─▶ slide parts (in sldIdLst order)
                                      ─▶ theme, slideMaster, slideLayouts
each slide: walkTree(spTree) ─▶ sp / pic / grpSp / graphicFrame(table|chart) / AlternateContent
```

Key design points:

- **Namespace-agnostic helpers** (`kid`, `kids`, `attr`) match by *local* element
  name, so the engine tolerates whatever namespace prefixes a real file uses.
- **Injectable XML parser** (`setDOMParser`/`setXMLSerializer`) — the browser
  uses the native `DOMParser`; the Node smoke suite injects `@xmldom/xmldom`.
  That's what lets the same parser code be unit-tested headless.
- **Placeholder inheritance:** a shape without its own geometry/text style
  resolves through `slideLayout → slideMaster`. Text size in particular follows a
  full chain (run `sz` ← paragraph `defRPr` ← shape `lstStyle` per level ← layout
  placeholder `lstStyle` ← master `txStyles`). This is why opening a real deck
  shows the *right* font sizes instead of defaults.
- **Modern pictures** (SVG icons, 3D, effects) are wrapped in
  `mc:AlternateContent`; `walkTree` recurses the `Fallback` (else `Choice`)
  branch, or images silently vanish.
- **Groups** flatten geometrically (child transforms baked into absolute
  coordinates) but membership is preserved by stamping a shared `groupId`.

## Writing — `buildPptx(pres, media)` / `exportPptxBlob(...)`

Builds a complete, valid package: `[Content_Types].xml`, `_rels`, presentation,
theme, slide master + layout, one part per slide, chart parts, notes, and media
files. The output opens without a repair prompt in PowerPoint.

Highlights:

- **Media de-duplication** keyed by `mediaId` (plus `#hex` for tinted SVG copies).
- **SVG dual-embed:** PowerPoint renders vectors only via the `asvg:svgBlip`
  extension and wants a PNG fallback blip beside it — the writer emits both, or
  the pure-graphic form (blip with no `r:embed`) for fallback-less SVG.
- **Groups** are re-wrapped into real `p:grpSp` with an identity child mapping
  (frame `chOff/chExt` = `off/ext`), so member coordinates pass through unchanged.

## Format gotchas (each one was a real bug)

These are encoded as guard assertions in the smoke suite so they can't regress:

- **Chart series names** (`c:tx`) accept only `strRef` or a literal `c:v` — a
  `strLit` there makes PowerPoint **repair-strip every chart**. Categories
  (`c:cat`) *do* allow `strLit`.
- **Office icon recolor:** icons carry CSS classes named `MsftOfcThm_*` that
  PowerPoint re-binds to the *live* theme color, overriding the literal CSS. A
  baked tint must also rename that prefix, or e.g. white reverts to accent1.
- **Line spacing** `spcPct` is a percentage of *single* spacing (~1.2em), not of
  font size — render lineHeight = `pct/100 * 1.2`.
- **Axis identity by `axPos`,** not element type: bar charts swap the category
  and value axes, so titles/styles must follow `axPos` (b/l), not catAx/valAx.
- **Theme sentinels** (`+mj-lt`/`+mn-lt`, scheme slots) are kept symbolic on
  parse — resolving them at parse time breaks theme propagation.

## Patterns — `pattern.ts`

`importPatternSlide(json, values?)` ingests SlideBazaar **pattern JSON**: raw
PresentationML fragments with `{{content.*}}` placeholders. It substitutes
placeholder values (schema-driven samples or caller overrides) and parses the
result like any slide. Used by `File → Import Pattern` and the dev hook
`window.__importPattern`.

## Security: hardening against untrusted input

Because the editor opens files fetched from object storage, `parsePptx` enforces
hard limits (`PARSE_LIMITS`, tunable) so a malicious or malformed `.pptx` can
only ever **fail loudly, never hang or OOM** the tab:

| Limit | Default | Guards against |
|---|---|---|
| `maxCompressedBytes` | 150 MB | oversized upload |
| `maxTotalUncompressedBytes` | 600 MB | **zip bombs** (also metered during extraction) |
| `maxEntryUncompressedBytes` | 200 MB | single-part bomb |
| `maxEntries` | 5000 | entry-count explosion |
| `maxSlides` | 1000 | slide flood |
| `maxGroupDepth` | 64 | deep-nesting stack overflow |

Additional posture (see [Code quality](12-code-quality.md#security)):

- **No XXE** — browser `DOMParser` doesn't resolve external entities or DTDs.
- **No `eval`/`new Function`** — the geometry evaluator is a hand-written interpreter.
- **No `innerHTML`/`dangerouslySetInnerHTML`** — raw file XML never reaches the live DOM.
- **SVG-as-image** — rendered via `<image href>` (non-scripting context), so embedded scripts don't run.
- **No SSRF** — `TargetMode="External"` relationship targets are never fetched.
- **Prototype-pollution guards** on file-derived dynamic keys (e.g. `avLst` guide names).

---

Next: [Rendering →](05-rendering.md)
