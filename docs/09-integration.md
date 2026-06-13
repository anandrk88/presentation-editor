# 09 · Integration & embedding

[← Index](../CLAUDE.md) · Source: [`src/util/embed.ts`](../src/util/embed.ts) · Full guide: [INTEGRATION.md](../INTEGRATION.md)

The editor is designed to be embedded in a host application as an iframe: hand it
a `.pptx` URL, let the user edit, get the edited file back. This page is the
summary; the complete reference (deploy steps, R2/S3 signing snippets, CORS,
acceptance checklist) is in **[INTEGRATION.md](../INTEGRATION.md)**.

## Two integration modes

### 1. URL parameters (simplest — best for R2/S3)
```html
<iframe src="https://editor.example.com/?file=<ENC_GET_URL>&title=Deck&saveUrl=<ENC_PUT_URL>"></iframe>
```
- `file` — signed GET URL of the `.pptx` to open (fetched client-side; needs CORS).
- `saveUrl` — presigned PUT URL; **Save** uploads the edited file there.
- `title` — header + export filename.

### 2. postMessage bridge (programmatic control)
Add `&embed=1&parentOrigin=<host origin>` and drive it:
```js
editor.contentWindow.postMessage({ type: "pe:load", url, title }, EDITOR_ORIGIN);
editor.contentWindow.postMessage({ type: "pe:save" }, EDITOR_ORIGIN);
```
Events back to the host (all tagged `source:"presentation-editor"`):
`pe:ready` (with `version`), `pe:loaded {title, slideCount}`, `pe:dirty {dirty}`,
`pe:document {data: ArrayBuffer, fileName}`, `pe:saved {via}`,
`pe:selection {selection}`, `pe:slide {slide}`, `pe:error {message}`.

### 3. Scripting API (read & change the document)
Inspect and mutate the document from code — active slide, selection, and each
element's text/image/fill/geometry, plus `setText`/`setImage`/`setFillColor`/
`setElementProperties`/`select`/`delete`/`undo`. Same-origin hosts call
`window.presentationEditor` directly; cross-origin hosts call any method by name
over `pe:invoke {requestId, method, args}` → `pe:result {requestId, ok, value}`.
**Full reference: [13 · Scripting API](13-scripting-api.md)** (source:
[`src/util/api.ts`](../src/util/api.ts)).

## Security of the bridge
With the bridge enabled, the editor accepts messages **only** from `parentOrigin`
and posts **only** to it. Without `parentOrigin` the bridge stays disabled. Only
`http(s)` URLs are accepted for `file`/`saveUrl`. Serve the editor from its **own
origin/subdomain** so an opened (untrusted) document can't reach the host app's
cookies/storage.

## Where it lives
`src/util/embed.ts` parses the URL config (`embedConfig()`), wires the bridge
(`initEmbedBridge`), and provides `uploadTo`/`notifyHost`. `App.tsx` bootstraps
it on mount and routes Save through it.

## Migrating from a server-based document editor
If you're replacing a Document-Server-style editor: its `document.url` →
`?file=`, its server `callbackUrl` → `?saveUrl=` (browser PUT) or the
`pe:document` event, its `onDocumentStateChange` → `pe:dirty`. The mapping table
is in [INTEGRATION.md](../INTEGRATION.md). Key difference: this editor saves
**from the browser** on user action, not server-to-server.

---

Next: [Build & release →](10-build-release.md)
