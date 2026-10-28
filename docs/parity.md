# Parity with the React adapter

The reference is `docx-editor`'s React package, version 2.33.0. This file records what the
Angular adapter reproduces, how, and what it does not — so "missing" is a documented decision
rather than an omission nobody noticed.

Legend: **yes** — implemented and exercised by `scripts/verify-demo.mjs`; **partial** — a
simpler rendering of the same capability; **no** — not implemented, and why.

## Shell

| React | Angular | Status |
| --- | --- | --- |
| `<DocxEditor document fonts …>` | `<docx-editor [document] [fonts] …>` | yes |
| `DocxEditor.Root` (provider-first) | `provideDocxEditor(options)` | yes |
| `DocxEditor.Viewport` | `<docx-editor-viewport>` | yes |
| `DocxEditor.Content` | `<docx-editor-content>` | yes |
| Title bar + editable document name | shell title bar, `[onTitleChange]` | yes |
| `DocxEditor.Toolbar` | `<docx-toolbar>` | yes |
| `DocxEditor.Menu` | `<docx-menu-bar>` | yes |
| `DocxEditor.Navigation` (headings) | `<docx-outline>` | partial — Headings tab only |
| `DocxEditor.HorizontalRuler` / `VerticalRuler` | — | no — planned; margins/indent drags need `dragIndent` |
| `DocxEditor.PageNumber` | `<docx-page-indicator>` | yes |
| `DocxEditor.Loading` | shell overlay | yes |
| `DocxEditor.OpenFailure` | shell dialog | yes |
| `DocxEditor.HyperLink` | — | no — planned |
| `DocxEditor.ContextMenu` | — | no — planned |
| `DocxEditor.ContentControl` | — | no — the engine slots exist; no inspector chrome |
| `DocxEditorUnsupportedElement` | — | no — custom rendering of unsupported content |
| `DocxEditor.ColorByChangeType` / `AuthorStyle` | `setRevisionStyles` via the store | partial — imperative only |
| Review rail (`@docx-editor.dev/pro`) | — | no — upstream module is separately licensed |

## Toolbar controls

Driven by `formattingBarChromeGroups()`; a slot renders only if the adapter has chrome for it.

| Slot | React | Angular | Notes |
| --- | --- | --- | --- |
| `history.undo` / `redo` | yes | yes | |
| `zoom.level` | stepper + menu | stepper + menu | Engine-owned scale. |
| `styles.style` | style picker with previews | yes | `getDocumentStyles()`. |
| `font.family` | searchable picker | yes | `getAvailableFonts()`. |
| `font.size` | stepper + presets + typing | stepper + presets | No free-typed size yet. |
| `text.bold` / `italic` / `underline` / `strike` | yes | yes | |
| `text.color` / `text.highlight` | split + Word picker | yes | Theme matrix, standard row, custom hex, clear. |
| `script.super` / `sub` | yes | yes | |
| `alignment.*` | yes | yes | |
| `direction.ltr` / `rtl` | yes | yes | |
| `list.bullet` / `numbered` / `indent` / `outdent` | yes | yes | |
| `list.lineSpacing` | menu | yes | `lineSpacingMenuRows()`. |
| `format.clear` / `format.painter` | yes | yes | |
| `review.editingMode` | pill + hints | yes | Refusals surfaced per row. |
| `image.insert` | file picker + decode | yes | Natural size via `createImageBitmap`. |
| `image.properties` / `wrap` / `altText` | dialogs | no | `ADAPTER_PENDING_SLOTS`. |
| `table.*` chrome | full table toolbar | no | Contextual group not rendered. |
| `contentControl.*` | inspector + toggles | no | |
| `review.addComment` / `authors` / `revisionMarkup` | yes | no | Needs the review rail / dialogs. |
| Toolbar overflow (`⋯`) | measured collapse | wraps | See roadmap. |

## Menu bar

`CHROME_MENUS` is rendered as data. Rows the adapter cannot complete are omitted:

| Row | Status |
| --- | --- |
| `file.open`, `file.save` | yes — host handler or packaged file picker / download |
| `file.exportMarkdown`, `file.exportPdf` | yes when `onExport` (or engine export handlers) is supplied |
| `file.print` | yes — host handler, else PDF print when a PDF handler exists, else `window.print()` |
| `file.pageSetup` | no — needs the dialog |
| `format.*` rows | yes |
| `insert.image`, `insert.textBox`, `insert.table` (grid), `insert.footnote`, `insert.endnote`, breaks, `insert.toc` | yes |
| `insert.watermark`, `insert.bookmark`, `insert.goToBookmark` | no — need dialogs/chrome |
| `review.*` rows | rendered; enabled state is the engine's (they need the review module) |
| `help` | empty, as upstream's default |

## Hooks → signals

| React hook | Angular |
| --- | --- |
| `useDocxEditor()` | `inject(DocxEditorStore)` |
| `useEditorState(selector, isEqual)` | `store.snapshot()`, `store.page`, `store.formatting`, … (`computed`/`memoize`) |
| `useEditorCommand(command)` | `store.commandState(slot)` or `store.can(command)` |
| `useEditorEvent(event, handler)` | `store.instance()?.on(event, handler)` |
| `useZoom()` | `store.zoom`, `store.setZoom`, `store.zoomIn/out` |
| `useDocxEditorRefApi` / `DocxEditorRef` | the component's public methods (`exportAs="docxEditor"`) |
| `useParagraphStyle` | `store.documentStyles()`, `store.formatting()?.styleId` |
| `useFontFamily` | `store.availableFonts()`, `store.formatting()?.fontFamily` |
| `useHistoryGroup` | `Editor.beginHistoryGroup()` on the facade |
| `useTranslation` / `LocaleProvider` | `provideDocxDefaults({ i18n })`, `[i18n]`, `store.t()` |
| `useChromeDir` | `store.chromeDir()` |
| `useDocumentOutline` | `store.outline()` |
| `useDocumentSearch` | — (engine `createDocumentSearch`; not wired) |
| `useReviewAuthors` | — (pro module) |
| `useHeaderFooterState` / `useNoteScopeState` | — (engine exposes them; no Angular chrome) |
| `usePageSetup` / `useParagraphFormat` | — (dialogs not ported) |

## Deliberate differences

1. **DI instead of projection.** React's context reaches projected children; Angular's does not,
   so the provider-first API is `provideDocxEditor()` on the host component.
2. **The bar wraps.** No overflow measurement yet. Upstream's stylesheet supports both (`flex-wrap`
   when it is not self-measuring, `data-overflow` when it is).
3. **One store, one editor.** `DocxEditorStore` is per-host-provider; nesting two editors means
   two providers, which is the intended usage.
4. **No StrictMode double-mount equivalent.** React creates and destroys an instance per effect
   run; Angular's `DestroyRef` gives one instance per host component.
