# 10 · Build & release

[← Index](../CLAUDE.md) · Source: [`scripts/`](../scripts/), [`.github/workflows/deploy.yml`](../.github/workflows/deploy.yml), [`vite.config.ts`](../vite.config.ts)

## Scripts (`package.json`)

| Command | What it does |
|---|---|
| `npm run dev` | Vite dev server at `:5173` |
| `npm run build` | `tsc` typecheck → Vite build → `dist/` → stamps `dist/version.json` |
| `npm run smoke` | Node round-trip engine test (~270 assertions) — see [Testing](11-testing.md) |
| `npm run fonts -- "<dir>"` | self-host fonts from a folder (see below) |
| `npm run release` | `npm version patch` → push commit + tag (triggers CI deploy) |

## Versioning

`vite.config.ts` injects `__APP_VERSION__` from `package.json` at build time
(declared in `src/global.d.ts`). It's shown in the About status and sent in the
`pe:ready` event. Each build also writes `dist/version.json`
(`{ version, commit, builtAt }`) so a host app can poll for new releases.

`vite.config.ts` uses `base: "./"` so the same build works at a domain root or
under a subpath.

## CI / deploy (`.github/workflows/deploy.yml`)

Every push to `main` (and tags `v*`): `npm ci` → `tsc --noEmit` → `npm run smoke`
→ `npm run build` → deploy. A broken engine never ships.

- **Default:** GitHub Pages (zero-config for testing).
- **Production options** (behind `if: false` + secrets): Cloudflare Pages, or
  S3 + CloudFront with cache-immutable hashed assets and no-cache entry points.

## How updates reach users

Vite emits **content-hashed assets**; `index.html` and `version.json` are
served no-cache. So **every fresh iframe load gets the new version
automatically**; open sessions keep their version until reloaded (safe — code is
never swapped under an open document). A host can poll `/version.json` (or read
`pe:ready.version`) to prompt long sessions to refresh. Full release/update flow
in [INTEGRATION.md §8](../INTEGRATION.md).

## <a id="fonts"></a>Self-hosted fonts (`scripts/install-fonts.mjs`)

```bash
npm run fonts -- "C:\path\to\fonts"   # then: npm run build
```
For each base family it copies the four styles the editor can pick
(Regular/Bold/Italic/BoldItalic = weight 400/700 × normal/italic) into
`public/fonts/`, generates `public/fonts/fonts.css` (`@font-face`,
`font-display: swap`), and writes `src/fonts/bundled.ts` (the family list merged
into the picker). `index.html` links the stylesheet. Width variants (Condensed,
Expanded, …) are skipped unless you pass `--all-widths`.

Fonts are referenced by name in the `.pptx` (PowerPoint uses the viewer's own
copy); bundling here governs **in-editor** rendering. Mind font licensing before
redistributing proprietary faces.

## Other tooling
- `scripts/gen-presets.mjs` — regenerates `src/render/presetDefs.json` (the 187
  preset geometry tables) from the vendored spec XML.
- `scripts/write-version.mjs` — emits `dist/version.json`.
- `scripts/repro-tint.mjs` — diagnostic: dumps how a real file persists SVG icons.

---

Next: [Testing & verification →](11-testing.md)
