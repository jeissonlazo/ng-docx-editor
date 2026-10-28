/**
 * ng-docx-editor — the Angular adapter for `@docx-editor.dev`.
 *
 * A thin renderer over the engine's `Editor` contract: the core owns parsing, layout, paint
 * and editing; this package owns the Angular components, the DI plumbing and the chrome the
 * engine's registry describes. It holds no editing-engine state of its own.
 *
 * ## Composition
 *
 * Two equivalent paths, exactly like the React and Vue adapters:
 *
 * 1. **Batteries included** — `<docx-editor [document]="bytes" />`. Chrome, menu bar,
 *    toolbar, headings pane and a painted editable document, in one tag.
 * 2. **Provider-first** — `provideDocxEditor(options)` on a host component, then compose
 *    `<docx-toolbar />`, `<docx-editor-viewport>` + `<docx-editor-content />` and any other
 *    part yourself; every part reads the same `DocxEditorStore` through DI.
 *
 * ## Styles
 *
 * The chrome is styled by the engine's own stylesheet, so load it once globally:
 *
 * ```json
 * "styles": ["node_modules/ng-docx-editor/styles.css", "src/styles.css"]
 * ```
 *
 * @packageDocumentation
 */

// ── Composition ───────────────────────────────────────────────────────────────────
export { DocxEditorComponent } from './lib/components/docx-editor';
export { DocxEditorViewportComponent } from './lib/components/docx-editor-viewport';
export { DocxEditorContentComponent } from './lib/components/docx-editor-content';
export { DocxOutlineComponent } from './lib/components/docx-outline';
export { DocxPageIndicatorComponent } from './lib/components/docx-page-indicator';

// ── Chrome parts ──────────────────────────────────────────────────────────────────
export { DocxToolbarComponent } from './lib/components/toolbar/docx-toolbar';
export { DocxToolbarButtonComponent } from './lib/components/toolbar/docx-toolbar-button';
export { DocxFontFamilyComponent } from './lib/components/toolbar/docx-font-family';
export { DocxFontSizeComponent } from './lib/components/toolbar/docx-font-size';
export { DocxStylePickerComponent } from './lib/components/toolbar/docx-style-picker';
export { DocxColorSplitComponent } from './lib/components/toolbar/docx-color-split';
export { DocxLineSpacingComponent } from './lib/components/toolbar/docx-line-spacing';
export { DocxZoomStepperComponent } from './lib/components/toolbar/docx-zoom-stepper';
export { DocxModePillComponent } from './lib/components/toolbar/docx-mode-pill';
export { DocxMenuBarComponent } from './lib/components/menu/docx-menu-bar';
export { DocxTableGridComponent } from './lib/components/menu/docx-table-grid';
export { DocxIconComponent } from './lib/components/shared/docx-icon';
export { DocxMenuDirective } from './lib/components/shared/docx-menu.directive';

// ── State, DI and lifetime ───────────────────────────────────────────────────────
export { DocxEditorStore } from './lib/core/docx-editor-store';
export { provideDocxEditor, provideDocxDefaults } from './lib/core/provide-docx-editor';
export type { DocxEditorDefaults } from './lib/core/provide-docx-editor';
export { DOCX_EDITOR_DIR, DOCX_EDITOR_I18N, DOCX_EDITOR_SHORTCUTS } from './lib/core/tokens';
export type {
  DocxEditorOptions,
  DocxEditorRef,
  DocxEditor,
  DocxFonts,
  EditorMode,
} from './lib/core/types';

// ── Chrome vocabulary helpers ────────────────────────────────────────────────────
export { slotInfo, slotIconPaths, slotLabelKey, type SlotInfo } from './lib/core/chrome-registry';
export { controlTitle, refusalText, slotLabel, type DocxTranslate } from './lib/core/labels';
export { ADAPTER_PENDING_SLOTS, isAdapterPending } from './lib/core/adapter-pending';
export { createDocxI18n, DEFAULT_DOCX_I18N, type DocxI18n } from './lib/core/i18n';
export {
  ZOOM_LEVELS,
  stepZoomLevel,
  zoomLevelForShortcut,
  zoomMenuRows,
  sameDeclaredZoomMode,
  isFitZoomMode,
  type ZoomMenuRow,
} from './lib/core/zoom';

// ── The engine surface a host drives, re-exported so one import is enough ────────
export type {
  CanResult,
  DocAnchor,
  DocRange,
  DocumentChange,
  DocumentHandle,
  DocumentSource,
  Editor,
  EditorCommand,
  EditorExecOptions,
  EditorFontError,
  EditorScope,
  EditorSnapshot,
  ExecResult,
  PageSetup,
  RunFormatting,
  Unsubscribe,
  ZoomMode,
} from '@docx-editor.dev/core/contracts/editor';
export type {
  ChromeControl,
  ChromeGroup,
  ChromeMenu,
  ChromeMenuEntry,
  ChromeSlotId,
  ChromeExportFormat,
  ChromeExportHandlers,
  DocxEditorInstance,
  EditorModule,
  FontConfigurationFragment,
  FontResolver,
  FontShapingOptions,
  ImageDecodePort,
  ResolvedRevisionMarkup,
  ReviewDisplayMode,
  RevisionMarkupOptions,
  ToolbarCommandState,
} from '@docx-editor.dev/core/editor';
export {
  CHROME_GROUPS,
  CHROME_MENUS,
  chromeMenuSlots,
  chromeSlotIsToggle,
  commandForSlot,
  commandForSlotValue,
  defaultChromeGroups,
  formattingBarChromeGroups,
  runToolbarCommand,
  toolbarCommandState,
} from '@docx-editor.dev/core/editor';
export type { TranslationKey, Translations } from '@docx-editor.dev/i18n';
