# ng-docx-editor

An Angular DOCX editor: Word-faithful paginated pages, a formatting bar, a menu bar and a
headings pane — the Angular adapter for the [`@docx-editor.dev`](https://docx-editor.dev)
engine, built from the same chrome registry and the same stylesheet as its React and Vue
adapters.

```html
<docx-editor [document]="bytes" title="Contract.docx" (change)="dirty.set(true)" />
```

## Install

```bash
npm install ng-docx-editor @docx-editor.dev/core @docx-editor.dev/i18n
npm install @docx-editor.dev/fonts   # optional: shaped, Word-accurate line breaks
```

## Styles (required)

```json
"styles": ["node_modules/ng-docx-editor/styles.css", "src/styles.css"]
```

## Font assets (recommended)

```json
"assets": [
  { "glob": "*.ttf", "input": "node_modules/@docx-editor.dev/fonts/assets", "output": "assets" },
  { "glob": "harfbuzz.wasm", "input": "node_modules/@docx-editor.dev/core/dist", "output": "." }
]
```

## Two composition paths

```ts
// 1. Batteries included.
import { DocxEditorComponent } from 'ng-docx-editor';

// 2. Provider-first: own the editor at your component's injector and compose the parts.
import { provideDocxEditor, DocxEditorStore, DocxToolbarComponent } from 'ng-docx-editor';

@Component({
  providers: [provideDocxEditor({ document: bytes })],
  imports: [DocxToolbarComponent, /* viewport, content, menu, outline … */],
})
```

## Public surface

- Components: `DocxEditorComponent`, `DocxEditorViewportComponent`, `DocxEditorContentComponent`,
  `DocxToolbarComponent`, `DocxMenuBarComponent`, `DocxOutlineComponent`,
  `DocxPageIndicatorComponent`, plus the individual toolbar parts (`DocxToolbarButtonComponent`,
  `DocxFontFamilyComponent`, `DocxFontSizeComponent`, `DocxStylePickerComponent`,
  `DocxColorSplitComponent`, `DocxLineSpacingComponent`, `DocxZoomStepperComponent`,
  `DocxModePillComponent`, `DocxTableGridComponent`, `DocxIconComponent`) and the `docxMenu`
  popover directive.
- State: `DocxEditorStore` (signals over `Editor.snapshot()`, engine-backed command state,
  host I/O helpers).
- Providers and tokens: `provideDocxEditor`, `provideDocxDefaults`, `DOCX_EDITOR_I18N`,
  `DOCX_EDITOR_DIR`, `DOCX_EDITOR_SHORTCUTS`.
- Chrome vocabulary: `slotInfo`, `slotIconPaths`, `slotLabelKey`, `controlTitle`,
  `refusalText`, `ADAPTER_PENDING_SLOTS`, and the engine's `CHROME_GROUPS` / `CHROME_MENUS` /
  `toolbarCommandState` re-exported so one import is enough.

The full guide — options, the imperative handle, i18n, dark mode, architecture and the parity
matrix against React — lives in the repository README:
<https://github.com/jeissonlazo/ng-docx-editor>.

## License

Apache-2.0. A derivative work of [docx-editor](https://github.com/eigenpal/docx-editor)
(Copyright 2026 EigenPal Inc.); see the repository `NOTICE`.
