import type {
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
  ZoomMode,
} from '@docx-editor.dev/core/contracts/editor';
import type {
  ChromeExportFormat,
  ChromeExportHandlers,
  ContentLockPolicy,
  DocxEditorInstance,
  EditorModule,
  FieldResultsMode,
  FontConfigurationFragment,
  FontResolver,
  FontShapingOptions,
  ImageDecodePort,
  LockedControlNavigation,
  ResolvedRevisionMarkup,
  ReviewDisplayMode,
  RevisionMarkupOptions,
} from '@docx-editor.dev/core/editor';
import type { FontConfiguration } from '@docx-editor.dev/core/contracts/editor';
import type { Translations } from '@docx-editor.dev/i18n';

/** The host editing mode, matching the toolbar's three-state pill. */
export type EditorMode = 'edit' | 'view' | 'suggesting';

/** Font sources: a full configuration, a fragment composed with defaults, or a resolver. */
export type DocxFonts = FontConfiguration | FontConfigurationFragment | FontResolver;

/**
 * Everything one editor can be told. A subset of the engine's `DocxEditorConfig` (the fields
 * a host sets) plus the host-owned actions the Angular chrome dispatches: reading bytes,
 * delivering bytes, exporting, printing.
 *
 * All fields are optional, and every field is applied the way the engine documents it: a
 * document identity change remounts; `fonts`, `author`, `locale`, `mode`, `translate`, `zoom`
 * and `zoomMode` are applied through instance setters so edits, selection and undo history
 * survive them; `modules` is sampled at mount.
 */
export interface DocxEditorOptions {
  /** A document to load: DOCX bytes, `'blank'`, or an existing handle. */
  document?: DocumentSource;
  /** Font sources for Word-accurate shaping and pagination. */
  fonts?: DocxFonts;
  /** Explicit shaping byte ceiling; a changed `maxFontBytes` remounts. */
  fontShaping?: FontShapingOptions;
  /** Image decoding port; a changed identity remounts. */
  imageDecodePort?: ImageDecodePort;
  /** Author for comments, replies and tracked changes. */
  author?: string;
  /** BCP-47 locale for regional date input and generated labels. */
  locale?: string;
  /** Capability modules (the pro review module, custom nodes). Sampled at mount. */
  modules?: readonly EditorModule[];
  /** The host editing mode. Omitted, the document's `w:trackRevisions` decides. */
  mode?: EditorMode;
  /** A fixed display scale. */
  zoom?: number;
  /** Where the display scale comes from. Defaults to the engine's `'auto'`. */
  zoomMode?: ZoomMode | 'auto';
  /** Controlled viewer markup settings. */
  revisionMarkup?: RevisionMarkupOptions;
  /** Called with settings proposed by an API or dialog change. */
  onRevisionMarkupChange?: (settings: ResolvedRevisionMarkup) => void;
  /** Initial revision projection. */
  reviewDisplayMode?: ReviewDisplayMode;
  /** How the reader edits saved field results. Read once, at mount. */
  fieldResults?: FieldResultsMode;
  /** Opt-in relaxation of content-control locks. Read once, at mount. */
  contentLockPolicy?: ContentLockPolicy;
  /** How the caret moves through a locked inline content control. Read once, at mount. */
  lockedControlNavigation?: LockedControlNavigation;

  /** Chrome catalogue for this editor, merged onto the inherited one. */
  i18n?: Translations;
  /** Forces the chrome writing direction; omitted, the catalogue or the page decides. */
  dir?: 'ltr' | 'rtl';
  /** Host label resolver; wins over the catalogue. */
  translate?: (key: string, params?: Record<string, string | number>) => string;

  /** Document name shown in the chrome's title bar. */
  title?: string;
  /** Chrome colour mode; `'system'` follows the OS. */
  colorMode?: 'light' | 'dark' | 'system';

  /** Open handler for File › Open. Without it a packaged file picker loads the bytes. */
  onOpen?: () => void;
  /** Save handler for File › Save. Without it the bytes download under `title`. */
  onSave?: () => void;
  /** Base file name used by the packaged downloads (no extension). */
  fileName?: string;
  /** Export handler for File › Export. Without it the export rows are unavailable. */
  onExport?: (format: ChromeExportFormat) => void | Promise<void>;
  /** Engine-side export handlers, used when `onExport` is absent. */
  exportHandlers?: ChromeExportHandlers;
  /** Print handler. Without it the packaged print runs on the document PDF or `window.print`. */
  onPrint?: () => void | Promise<void>;

  /** Fired once per instance, after the first document has painted. */
  onReady?: (editor: Editor) => void;
  /** Fired when the document changes (revision + identity deltas, not bytes). */
  onChange?: (change: DocumentChange) => void;
  /** Fired with the typed font failure the engine rejected. */
  onFontError?: (error: EditorFontError) => void;
}

/**
 * The imperative handle, matching the React and Vue adapters member for member.
 *
 * Every member forwards to the `Editor` facade and is safe to call before the editor has
 * mounted: mutations no-op, reads return the honest empty answer, so a host can hold the
 * handle from the first render without guarding it.
 */
export interface DocxEditorRef {
  /** Load a document: DOCX bytes, `'blank'`, or an existing handle. */
  load(document: DocumentSource): void;
  /** Serialize the current document; `null` when no editor is mounted. */
  save(): Promise<ArrayBuffer | null>;
  /** Identity and revision of the loaded document; `null` before mount. */
  getDocumentHandle(): DocumentHandle | null;
  /** The full `Editor` facade for advanced callers; `null` before mount. */
  getEditor(): Editor | null;
  /** Put the caret in the pages. */
  focus(): void;
  /** Run a typed command through the facade. */
  exec(command: EditorCommand, options?: EditorExecOptions): ExecResult;
  /** The current read model; a loading, non-editable snapshot before mount. */
  snapshot(options?: { scope?: EditorScope }): EditorSnapshot;
}

/** The concrete facade type the store exposes. */
export type DocxEditor = DocxEditorInstance;
