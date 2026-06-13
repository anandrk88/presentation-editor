# 11 · Testing & verification

[← Index](../CLAUDE.md) · Source: [`scripts/smoke-pptx.mjs`](../scripts/smoke-pptx.mjs), [`scripts/smoke-entry.ts`](../scripts/smoke-entry.ts)

## The smoke round-trip suite

```bash
npm run smoke      # must end with "SMOKE OK"
```

The suite is the project's primary safety net (~270 assertions). It works by
**round-tripping the model through the real engine**:

```
build a model with a feature  ──buildPptx()──▶  .pptx bytes
        │                                              │
        │                                       parsePptx()
        ▼                                              ▼
    assert the re-parsed model still has the feature, and 0 import warnings
```

How it runs headless: `smoke-pptx.mjs` uses Vite to bundle the real
`parse`/`write`/`model` code (`smoke-entry.ts`) into a single ESM file, then
imports it in Node. The parser's injectable XML implementation
(`setDOMParser`/`setXMLSerializer`) is fed `@xmldom/xmldom`, so the exact
production engine code is exercised without a browser.

It also validates **package structure** (every required part present,
content-types cover every part, all relationship targets resolve) and **XML
well-formedness** of every part — so the output is provably a valid `.pptx`.

## What it covers

Shapes (preset geometry + adjust values, freeform custGeom), text (styles,
bullets, columns, super/sub, highlight, size inheritance through master/layout),
tables (grid, styles, merges, the style-option flags), charts (8 kinds, grouping,
scatter, radar, the formatting fields, per-element text styling, **and the
PowerPoint-repair-trigger guard**), images (SVG dual-embed, crop, corners,
multi-image AlternateContent), fills (gradient/picture/pattern, backgrounds),
groups (membership round-trip + identity child mapping), themes/fonts, notes,
line arrowheads + dashes, gridlines, pie slice colors, and the security limits.

Each historical format bug (see [OOXML engine](04-ooxml-engine.md#format-gotchas))
has a dedicated assertion so it can't regress.

## Live-browser verification

Functional/visual behavior (drags, crop handles, menus, rendering) is verified by
driving the running app in a real browser and asserting on store state + the DOM,
plus screenshots. The store is exposed as `window.__store` in dev for scripted
checks.

### Lessons baked into the methodology
- The app's `beforeunload` dirty-guard blocks `location.reload()`; stash + clear
  `dirty` before reloading a working tab.
- `setPointerCapture` retargets derived events — synthetic drag tests temporarily
  no-op it.
- Store edits (`store.ts`) force a full reload that wipes the in-progress
  document — prefer component edits while a user is editing live; never drive the
  user's live session for testing.

## What's missing

There are **no component/unit tests** beyond the engine round-trip, no a11y
tests, and no automated visual-regression baseline. The engine is very
well-covered; the UI is covered manually. See [Code quality](12-code-quality.md).

---

Next: [Code quality assessment →](12-code-quality.md)
