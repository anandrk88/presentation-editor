# 12 · Code quality assessment

[← Index](../CLAUDE.md)

An honest review of how the codebase is structured and how well it's written.
Written for a maintainer deciding what to trust and what to refactor.

## Summary

> **A well-architected, exceptionally well-tested OOXML engine wrapped in a
> capable but heavyweight UI.** The core (model + engine + render) is the kind of
> code you'd be happy to inherit. The UI layer works and is feature-rich, but has
> grown a few large files that are the main tech debt. It's a strong
> production-leaning codebase, not yet "big-team enterprise" (no a11y/i18n,
> shallow UI test coverage).

**Overall: B+ / A-.** The engine alone is an A; the UI pulls the average down.

## What's strong

**1. Clean layered architecture.** Five layers with strict downward
dependencies (model → engine → render → state → UI). You can reason about any
layer in isolation. The model is inert data; the engine and renderer are pure
functions of it. This is textbook and consistently followed.

**2. A pure, testable engine.** `parse.ts`/`write.ts` have no DOM/browser
coupling (the XML parser is injected), so the **real** engine code runs headless
in Node. That's why the ~270-assertion round-trip suite is trustworthy rather
than theater — it exercises production code paths, not mocks.

**3. Format fidelity earned the hard way.** The genuinely tricky parts of OOXML
are handled correctly and *documented at the call site*: the `c:tx` strLit
repair trap, SVG dual-embed, `MsftOfcThm_` theme rebinding, line-spacing
semantics, axis-identity-by-position, text-size inheritance chains. Each has a
regression guard. This is the most valuable and least-replaceable code in the
repo.

**4. Idiomatic, minimal state management.** One observable store +
`useSyncExternalStore`. No Redux/MobX/context maze. The whole app state is one
greppable object (`window.__store.state`), undo is correct-by-construction
(immutable snapshots), and the preview/commit pattern makes drags one undo step.

**5. Strong typing.** Discriminated unions for `Shape`, `Fill`, `ColorRef` give
exhaustive, safe handling; `tsc --noEmit` is clean and gates CI.

**6. Security-conscious for untrusted input.** Explicit parse limits, no
`eval`/`new Function`, no `innerHTML`, prototype-pollution guards, SVG rendered
as a non-scripting image. A deliberate review pass, not an afterthought.

**7. Performance was measured, not guessed.** Memoized thumbnails (commit
re-renders only the changed slide), snapshot undo, preview-path drags. A 60-slide
/ 600-shape deck loads in ~200 ms and edits within a frame — verified, not
assumed.

## What's tech debt

**1. A few oversized files.** The clearest issue:

| File | Lines | Problem |
|---|---|---|
| `RightPanel.tsx` | ~1340 | every settings pane + the fill editor + gradient slider + table/chart panes in one file |
| `parse.ts` | ~1340 | the whole reader; dense but cohesive |
| `EditorCanvas.tsx` | ~1070 | all pointer logic + every selection-chrome variant + rulers + context menu |
| `Ribbon.tsx` | ~990 | all tabs + galleries + dialogs |

`parse.ts` being large is defensible (one cohesive responsibility). The **UI**
monoliths are the real debt — `RightPanel` and `EditorCanvas` should be split
(one file per pane; extract the drag-kind handlers). They're navigable today
because of consistent structure and good comments, but they raise the cost of
change and the risk of merge conflicts.

**2. Shallow UI test coverage.** The engine is exhaustively tested; the
components and most store methods are verified only manually / via live-browser
scripting. There are no component tests, no isolated store-method unit tests, no
visual-regression baseline. For an embeddable product this is the gap I'd close
first.

**3. Snapshot undo deep-clones the presentation per commit.** Simple and
correct, but O(document size) memory per history entry (capped at 200). Fine for
normal decks; structural sharing (or a patch-based history) would scale better
for very large ones.

**4. Dev affordances leak toward production.** `window.__store` and
`window.__importPattern` are always attached. Harmless, but should be gated
behind `import.meta.env.DEV`.

**5. Styling is one 640-line `styles.css` + inline styles.** Fine at this scale,
but there's no CSS-module/scoping discipline, so class-name collisions are
prevented only by convention.

**6. Not yet "enterprise-grade" in the non-functional dimensions.** No
accessibility (ARIA roles, keyboard menu nav, focus management), no
internationalization, no telemetry/error reporting hooks. These are
green-field, not broken — but they're the distance to true enterprise software.

## Consistency & craft

- **Naming** is clear and consistent (`makeShape`, `resolveColor`,
  `tableCellStyle`, `chartPartRegions`).
- **Comments** are unusually good where it matters — they explain *why* (the
  format gotcha, the layering reason), not *what*. Dense modules (`parse.ts`,
  `write.ts`) carry their reasoning.
- **No dead code of note**; factories and helpers are reused rather than copied.
- **Error handling** is pragmatic: the parser degrades (warnings + placeholders)
  rather than throwing on imperfect files, which is right for a real-world
  importer.

## If you had three refactors to spend

1. **Split `RightPanel.tsx` and `EditorCanvas.tsx`** into per-pane / per-drag-kind
   modules. Biggest maintainability win, low risk.
2. **Add component + store-method tests** (Vitest + Testing Library) to match the
   engine's rigor. Biggest confidence win for the UI.
3. **Gate dev globals** and add an error-reporting hook for embedded deployments.

## Bottom line

The hard part of a `.pptx` editor — a faithful, bidirectional OOXML engine — is
done very well and is well-protected by tests. The UI is feature-complete and
thoughtfully built but carries the typical "fast-moving feature work" debt of a
few large files and thin automated UI coverage. Nothing here is alarming;
everything flagged is a known, bounded refactor rather than a structural flaw.

---

[← Back to index](../CLAUDE.md)
