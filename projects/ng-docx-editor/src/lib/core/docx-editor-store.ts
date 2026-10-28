import { DestroyRef, Injectable, computed, inject, signal } from '@angular/core';
import { localizeDisabledReason } from '@docx-editor.dev/i18n';
import type { TranslationKey } from '@docx-editor.dev/i18n';
import {
  chromeMenuSlots,
  LOADING_SNAPSHOT,
  resolveRevisionMarkup,
  runChromeExport,
  runChromePrint,
  runToolbarCommand,
  toolbarCommandState,
} from '@docx-editor.dev/core/editor';
import type {
  ChromeExportFormat,
  ChromeSlotId,
  DocxEditorInstance,
  ToolbarCommandState,
} from '@docx-editor.dev/core/editor';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import type {
  DocumentChange,
  DocumentSource,
  Editor,
  EditorCommand,
  EditorExecOptions,
  EditorScope,
  EditorSnapshot,
  ExecResult,
  ZoomMode,
} from '@docx-editor.dev/core/contracts/editor';
import { DOCX_EDITOR_DIR, DOCX_EDITOR_I18N } from './tokens';
import { createDocxI18n, DEFAULT_DOCX_I18N, type DocxI18n } from './i18n';
import { deferredTick } from './deferred-notifier';
import { sameDeclaredZoomMode, stepZoomLevel, zoomLevelForShortcut } from './zoom';
import type { DocxEditorOptions, DocxEditorRef, EditorMode } from './types';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** The slots the host, not the engine, dispatches. */
const HOST_SLOTS: ReadonlySet<string> = new Set<ChromeSlotId>([
  'file.open',
  'file.save',
  'file.print',
  'file.exportMarkdown',
  'file.exportPdf',
]);

/**
 * One editor's state, lifetime and imperative surface — the Angular twin of React's
 * `DocxEditor.Root` + `useDocxEditor()` + the ref bridge, in a single injectable.
 *
 * The store is provided per host (`provideDocxEditor()`), or by `<docx-editor>` itself, so a
 * toolbar, an outline pane and the mount point all read the same instance through DI — the
 * Angular equivalent of the React context the packaged chrome consumes.
 *
 * State is published through signals: `version` ticks once per coalesced engine event and
 * every derived signal (`snapshot`, `outline`, `commandState`) reads it, so a control far from
 * the pages re-renders when the engine moves and sleeps through unrelated changes.
 */
@Injectable()
export class DocxEditorStore implements DocxEditorRef {
  private readonly destroyRef = inject(DestroyRef);
  private readonly globalI18n = inject(DOCX_EDITOR_I18N, { optional: true });
  private readonly globalDir = inject(DOCX_EDITOR_DIR, { optional: true });

  /** The options in force. */
  private readonly optionsSignal = signal<DocxEditorOptions>({});
  readonly options = this.optionsSignal.asReadonly();

  private readonly instanceSignal = signal<DocxEditorInstance | null>(null);
  /** The live facade, or `null` before the first mount and after `destroy()`. */
  readonly instance = this.instanceSignal.asReadonly();

  /** Bumps once per coalesced engine event. Everything derived reads it. */
  readonly version = signal(0);

  /** A user-facing notice (a refusal, a failed export); the shell renders and dismisses it. */
  readonly notice = signal<string | null>(null);

  /** Set when the platform reports a dark colour scheme. */
  private readonly systemDark = signal(false);

  private mount: HTMLElement | null = null;
  private unsubscribes: (() => void)[] = [];
  private appliedZoom: number | undefined;
  private appliedZoomMode: ZoomMode | 'auto' | undefined;
  private appliedMode: EditorMode | undefined;
  private applyingMarkup = false;
  private readyFired = false;
  private destroyed = false;
  private memoVersion = -1;
  private readonly memo = new Map<string, unknown>();

  private readonly notify = deferredTick(() => {
    if (this.destroyed) return;
    this.version.update((value) => value + 1);
  });

  constructor() {
    this.destroyRef.onDestroy(() => this.destroy());
    if (typeof window !== 'undefined' && typeof window.matchMedia === 'function') {
      const query = window.matchMedia('(prefers-color-scheme: dark)');
      this.systemDark.set(query.matches);
      const listener = (event: MediaQueryListEvent) => this.systemDark.set(event.matches);
      query.addEventListener?.('change', listener);
      this.destroyRef.onDestroy(() => query.removeEventListener?.('change', listener));
    }
  }

  // ── Configuration ────────────────────────────────────────────────────────────────

  /**
   * Merge `patch` into the options and apply it.
   *
   * Idempotent: a patch whose values are all already in force is a NO-OP — it does not write
   * the options signal at all. That matters because a host applies its inputs from an
   * `effect()`: writing an unconditionally fresh options object there would re-trigger that
   * same effect forever. A document identity change remounts the editor (the engine has no
   * in-place way to swap a session); everything else goes through the instance setters so
   * edits, caret and undo history survive.
   */
  configure(patch: DocxEditorOptions): void {
    if (this.destroyed) return;
    const previous = this.optionsSignal();
    if (shallowEqualOptions(previous, patch)) return;
    const next: DocxEditorOptions = { ...previous, ...patch };
    this.optionsSignal.set(next);
    const instance = this.instanceSignal();
    if (!instance) {
      this.create(next);
      return;
    }
    if (this.mustRemount(previous, next)) {
      this.remount(next);
      return;
    }
    this.applyLive(previous, next, instance);
  }

  /** Resolve an i18n state for one set of options. */
  private i18nFor(options: DocxEditorOptions): DocxI18n {
    const base = this.globalI18n
      ? createDocxI18n(DEFAULT_DOCX_I18N, this.globalI18n, undefined)
      : DEFAULT_DOCX_I18N;
    return createDocxI18n(base, options.i18n, this.globalDir ?? options.dir);
  }

  /** The i18n state the current options resolve to. */
  readonly i18n = computed<DocxI18n>(() => this.i18nFor(this.optionsSignal()));

  /** The label resolver every packaged part uses: the host's, else the catalogue's. */
  readonly t = computed<(key: string, params?: Record<string, string | number>) => string>(() => {
    const options = this.optionsSignal();
    const i18n = this.i18n();
    const host = options.translate;
    if (host) return host;
    return (key: string, params?: Record<string, string | number>) =>
      i18n.t(key as TranslationKey, params);
  });

  /** The chrome direction: a decided one, else `'ltr'` as the shipped catalogues are. */
  readonly chromeDir = computed<'ltr' | 'rtl'>(() => this.i18n().dir);

  /** Whether the chrome paints dark, from `colorMode` and the OS setting. */
  readonly isDark = computed(() => {
    const mode = this.optionsSignal().colorMode ?? 'light';
    return mode === 'system' ? this.systemDark() : mode === 'dark';
  });

  private mustRemount(previous: DocxEditorOptions, next: DocxEditorOptions): boolean {
    if (previous.document !== next.document) return true;
    if (previous.fontShaping?.maxFontBytes !== next.fontShaping?.maxFontBytes) return true;
    if (previous.imageDecodePort !== next.imageDecodePort) return true;
    return false;
  }

  private applyLive(
    previous: DocxEditorOptions,
    next: DocxEditorOptions,
    instance: DocxEditorInstance
  ): void {
    if (previous.fonts !== next.fonts) {
      void instance.setFonts(next.fonts).catch(() => {
        // Font failures report through `onFontError`; the editor stays on the fallback.
      });
    }
    if (previous.i18n !== next.i18n || previous.translate !== next.translate) {
      instance.setTranslate(this.t());
    }
    if (previous.author !== next.author) instance.setAuthor(next.author);
    if (previous.locale !== next.locale) instance.setLocale(next.locale);
    if (previous.mode !== next.mode) {
      instance.setMode(next.mode);
      this.appliedMode = next.mode;
    }
    if (previous.revisionMarkup !== next.revisionMarkup && next.revisionMarkup !== undefined) {
      this.applyRevisionMarkup(next.revisionMarkup);
    }
    this.applyZoom(instance, next);
  }

  private applyRevisionMarkup(settings: NonNullable<DocxEditorOptions['revisionMarkup']>): void {
    const instance = this.instanceSignal();
    if (!instance) return;
    const resolved = resolveRevisionMarkup(settings);
    if (JSON.stringify(resolved) === JSON.stringify(instance.snapshot().revisionMarkup)) return;
    this.applyingMarkup = true;
    try {
      instance.setRevisionMarkup(resolved);
    } finally {
      this.applyingMarkup = false;
    }
  }

  /**
   * Apply `zoom` and `zoomMode` with the engine's own precedence rules: the level first, then
   * a declared fit — `setZoom` leaves any fit by design, so a host declaring both gets its
   * mode back rather than whichever call happened to run last.
   */
  private applyZoom(instance: DocxEditorInstance, options: DocxEditorOptions): void {
    let zoomChanged = false;
    if (options.zoom !== undefined && options.zoom !== this.appliedZoom) {
      this.appliedZoom = options.zoom;
      instance.setZoom(options.zoom);
      zoomChanged = true;
    }
    const modeMoved =
      options.zoomMode !== undefined &&
      (this.appliedZoomMode === undefined ||
        !sameDeclaredZoomMode(this.appliedZoomMode, options.zoomMode));
    const declaredFit =
      options.zoomMode !== undefined &&
      (options.zoomMode === 'auto' ||
        (typeof options.zoomMode === 'object' && options.zoomMode.type === 'fit'));
    if (modeMoved || (zoomChanged && declaredFit)) {
      this.appliedZoomMode = options.zoomMode;
      instance.setZoomMode(options.zoomMode!);
    }
  }

  // ── Lifetime ─────────────────────────────────────────────────────────────────────

  private create(options: DocxEditorOptions): void {
    const translate = this.t();
    let instance: DocxEditorInstance;
    try {
      instance = createDocxEditor({
        ...(options.document !== undefined ? { document: options.document } : {}),
        ...(options.fonts ? { fonts: options.fonts } : {}),
        ...(options.fontShaping ? { fontShaping: options.fontShaping } : {}),
        ...(options.author !== undefined ? { author: options.author } : {}),
        ...(options.locale !== undefined ? { locale: options.locale } : {}),
        translate,
        ...(options.mode !== undefined ? { mode: options.mode } : {}),
        ...(options.modules !== undefined ? { modules: options.modules } : {}),
        ...(options.revisionMarkup !== undefined ? { revisionMarkup: options.revisionMarkup } : {}),
        ...(options.reviewDisplayMode !== undefined
          ? { reviewDisplayMode: options.reviewDisplayMode }
          : {}),
        ...(options.fieldResults !== undefined ? { fieldResults: options.fieldResults } : {}),
        ...(options.contentLockPolicy !== undefined
          ? { contentLockPolicy: options.contentLockPolicy }
          : {}),
        ...(options.lockedControlNavigation !== undefined
          ? { lockedControlNavigation: options.lockedControlNavigation }
          : {}),
        ...(options.zoom !== undefined ? { zoom: options.zoom } : {}),
        ...(options.zoomMode !== undefined ? { zoomMode: options.zoomMode } : {}),
        ...(options.imageDecodePort ? { imageDecodePort: options.imageDecodePort } : {}),
        onFontError: (error) => this.optionsSignal().onFontError?.(error),
      });
    } catch (error) {
      console.error('[ng-docx-editor] The editor could not be created.', error);
      this.notice.set(error instanceof Error ? error.message : String(error));
      return;
    }
    this.appliedZoom = options.zoom;
    this.appliedZoomMode = options.zoomMode;
    this.appliedMode = options.mode;
    this.readyFired = false;
    this.instanceSignal.set(instance);
    this.subscribe(instance);
    if (this.mount) this.safeAttach(instance, this.mount);
    this.fireReady(instance);
    this.notify();
  }

  private subscribe(instance: DocxEditorInstance): void {
    for (const off of this.unsubscribes) off();
    this.unsubscribes = [
      instance.on('change', (change: DocumentChange) => {
        this.optionsSignal().onChange?.(change);
        this.notify();
      }),
      instance.on('selectionChange', () => this.notify()),
      instance.on('error', () => this.notify()),
      instance.on('revisionMarkupChange', (settings) => {
        if (!this.applyingMarkup) {
          type Change = NonNullable<DocxEditorOptions['onRevisionMarkupChange']>;
          (this.optionsSignal().onRevisionMarkupChange as Change | undefined)?.(
            settings as Parameters<Change>[0]
          );
        }
        this.notify();
      }),
    ];
  }

  private remount(options: DocxEditorOptions): void {
    this.teardownInstance();
    this.create(options);
  }

  private teardownInstance(): void {
    for (const off of this.unsubscribes) off();
    this.unsubscribes = [];
    const instance = this.instanceSignal();
    this.instanceSignal.set(null);
    if (instance) {
      try {
        instance.detach();
      } catch {
        // Already gone.
      }
      instance.destroy();
    }
    this.memo.clear();
    this.memoVersion = -1;
  }

  private fireReady(instance: DocxEditorInstance): void {
    if (this.readyFired) return;
    const fire = () => {
      if (this.readyFired) return;
      this.readyFired = true;
      this.optionsSignal().onReady?.(instance as unknown as Editor);
    };
    if (!instance.snapshot().isOpening) {
      fire();
      return;
    }
    const off = instance.on('change', () => {
      off();
      fire();
    });
    this.unsubscribes.push(off);
  }

  /** Mount the editor in `element`. Called by `<docx-editor-content>`. */
  attach(element: HTMLElement): void {
    this.mount = element;
    const instance = this.instanceSignal();
    if (!instance) {
      this.create(this.optionsSignal());
      return;
    }
    this.safeAttach(instance, element);
  }

  private safeAttach(instance: DocxEditorInstance, element: HTMLElement): void {
    try {
      instance.attach(element);
    } catch (error) {
      // A crashed open is already a refusal on the snapshot, which the shell shows.
      if (instance.snapshot().openFailure?.code !== 'open.failed') throw error;
      console.error('[ng-docx-editor] The document failed to open.', error);
    }
  }

  /** Park the editor, keeping its history, selection and reading position. */
  detach(): void {
    const instance = this.instanceSignal();
    if (!instance) return;
    try {
      instance.detach();
    } catch {
      // Detaching twice is a documented no-op.
    }
  }

  /** Release the instance and every subscription. Terminal. */
  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.mount = null;
    this.teardownInstance();
  }

  // ── Read model ───────────────────────────────────────────────────────────────────

  /**
   * The engine's read model, or the frozen loading snapshot before the first mount.
   *
   * Reading it inside a template (or a computed) subscribes to every state tick, and the
   * facade's own version cache means repeated reads between ticks do no work.
   */
  snapshot(options?: { scope?: EditorScope }): EditorSnapshot {
    this.version();
    return this.instanceSignal()?.snapshot(options) ?? LOADING_SNAPSHOT;
  }

  /** Whether the document can be edited right now. */
  readonly editable = computed(() => this.snapshot().editable);
  /** One-based current page and the page count. */
  readonly page = computed(() => this.snapshot().page);
  /** Formatting at the selection, for the toolbar's value displays. */
  readonly formatting = computed(() => this.snapshot().formatting);
  /** The table at the selection, when there is one. */
  readonly table = computed(() => this.snapshot().table);
  /** Formatting at the selection, as the engine reads it (alias of `formatting`). */
  readonly selectionFormatting = computed(() => this.snapshot().formatting);
  /** Display scale; `1` is 100%. */
  readonly zoom = computed(() => this.snapshot().zoom);
  /** Whether the scale is fixed or tracks the viewport. */
  readonly zoomMode = computed(() => this.snapshot().zoomMode);
  /** Undo availability. */
  readonly canUndo = computed(() => this.snapshot().canUndo ?? false);
  /** Redo availability. */
  readonly canRedo = computed(() => this.snapshot().canRedo ?? false);
  /** Whether an open is in flight. */
  readonly isLoading = computed(() => this.snapshot().isLoading || this.snapshot().isOpening === true);
  /** Why the last open failed, when it did. */
  readonly openFailure = computed(() => this.snapshot().openFailure ?? null);
  /** The engine's editing mode. */
  readonly editingMode = computed(() => this.snapshot().editingMode ?? 'editing');
  /** Page marks visibility. */
  readonly showParagraphMarks = computed(() => this.snapshot().showParagraphMarks ?? false);
  /** Bookmark marks visibility. */
  readonly showBookmarks = computed(() => this.snapshot().showBookmarks ?? false);
  /** The review pane's resolved state, when a review module is installed. */
  readonly reviewPaneOpen = computed(() => this.snapshot().reviewPaneOpen ?? false);
  /**
   * Cache one derivation for the current version, so a template that reads it repeatedly (or
   * several templates reading it) costs one engine call per state tick.
   */
  private memoize<T>(key: string, compute: () => T): T {
    const version = this.version();
    if (version !== this.memoVersion) {
      this.memo.clear();
      this.memoVersion = version;
    }
    if (!this.memo.has(key)) this.memo.set(key, compute());
    return this.memo.get(key) as T;
  }

  /** The heading outline, in document order. */
  outline(): readonly { readonly text: string; readonly level: number; readonly blockId: string }[] {
    return this.memoize('outline', () => {
      const instance = this.instanceSignal();
      if (!instance) return [];
      try {
        return instance.getOutline();
      } catch {
        return [];
      }
    });
  }

  /** The paragraph styles the document defines, for the style picker. */
  documentStyles(): readonly {
    readonly styleId: string;
    readonly name: string;
    readonly type: string;
    readonly preview: {
      readonly fontFamily: string | null;
      readonly fontSizePt: number | null;
      readonly bold: boolean;
      readonly italic: boolean;
      readonly color: string | null;
    };
  }[] {
    return this.memoize('styles', () => {
      const instance = this.instanceSignal();
      if (!instance) return [];
      try {
        return instance.getDocumentStyles();
      } catch {
        return [];
      }
    });
  }

  /** Every family the font picker offers. */
  availableFonts(): readonly string[] {
    return this.memoize('fonts', () => {
      const instance = this.instanceSignal();
      if (!instance) return [];
      try {
        return instance.getAvailableFonts();
      } catch {
        return [];
      }
    });
  }

  /** The document theme's ten picker colours, six-digit hex without `#`. */
  themeColors(): readonly { readonly slot: string; readonly hex: string }[] {
    return this.memoize('themeColors', () => {
      const instance = this.instanceSignal();
      if (!instance) return [];
      try {
        return instance.getDocumentThemeColors();
      } catch {
        return [];
      }
    });
  }

  /** Enabled/active state for one chrome slot, from the engine, with host slots included. */
  commandState(slot: ChromeSlotId): ToolbarCommandState {
    return this.memoize(`state:${slot}`, () => {
      const instance = this.instanceSignal();
      if (HOST_SLOTS.has(slot)) {
        const handlers = this.hostHandler(slot);
        return {
          id: slot,
          enabled: handlers !== null,
          disabledReason: handlers === null ? 'not wired to an editor command' : null,
          active: false,
        };
      }
      return toolbarCommandState(instance, slot);
    });
  }

  /** Enabled/active state for every slot the packaged menu bar places. */
  menuCommandStates(): ReadonlyMap<ChromeSlotId, ToolbarCommandState> {
    return this.memoize('menuStates', () => {
      const states = new Map<ChromeSlotId, ToolbarCommandState>();
      for (const slot of chromeMenuSlots()) states.set(slot, this.commandState(slot));
      return states;
    });
  }

  // ── Commands ─────────────────────────────────────────────────────────────────────

  /** Run one chrome slot through `can` and then `exec`, exactly as the toolbar does. */
  run(slot: ChromeSlotId, value?: unknown): ExecResult {
    const host = this.hostHandler(slot);
    if (host && value === undefined) {
      void host();
      return { ok: true, changed: false };
    }
    const instance = this.instanceSignal();
    const result =
      value === undefined
        ? runToolbarCommand(instance, slot)
        : runToolbarCommand(instance, slot as never, value as never);
    if (!result.ok && result.reason) this.reportRefusal(result.reason);
    return result;
  }

  /** Whether the command is currently applied at the selection. */
  isActive(command: EditorCommand): boolean {
    const instance = this.instanceSignal();
    return instance ? instance.isActive(command) : false;
  }

  /** Ask the engine whether it would honour a command now. */
  can(command: EditorCommand, options?: EditorExecOptions) {
    const instance = this.instanceSignal();
    if (!instance) return { ok: false as const, code: 'unsupported' as const, reason: 'editor is not ready' };
    return instance.can(command, options);
  }

  /** Run a typed command. */
  exec(command: EditorCommand, options?: EditorExecOptions): ExecResult {
    const instance = this.instanceSignal();
    if (!instance) {
      return { ok: false, code: 'unsupported', reason: 'the editor is not mounted yet' };
    }
    const result = instance.exec(command, options);
    if (!result.ok && result.reason) this.reportRefusal(result.reason);
    return result;
  }

  /** Undo one step. */
  undo(): ExecResult {
    return this.run('history.undo');
  }

  /** Redo one step. */
  redo(): ExecResult {
    return this.run('history.redo');
  }

  /** Toggle a mark by name (`bold`, `italic`, `underline`, `strike`). */
  toggleMark(mark: string): ExecResult {
    return this.exec({ type: 'toggleMark', mark });
  }

  /** Set the paragraph style from the document's own styles. */
  setParagraphStyle(styleId: string): ExecResult {
    return this.run('styles.style', styleId);
  }

  /** Set the font family for the selection. */
  setFontFamily(family: string): ExecResult {
    return this.run('font.family', family);
  }

  /** Set the font size, in half-points (the engine's unit — `22` is 11pt). */
  setFontSize(halfPoints: number): ExecResult {
    return this.run('font.size', halfPoints);
  }

  /** Set the text colour, as six-digit hex without `#`. */
  setTextColor(hex: string): ExecResult {
    return this.run('text.color', hex);
  }

  /** Set the highlight colour; `null` clears it. */
  setHighlight(color: string | null): ExecResult {
    return this.run('text.highlight', color);
  }

  /** Set the editing mode (the toolbar's three-state pill). */
  setEditingMode(mode: 'editing' | 'suggesting' | 'viewing'): ExecResult {
    return this.run('review.editingMode', mode);
  }

  /** Insert a table of `rows` × `cols`. */
  insertTable(rows: number, cols: number): ExecResult {
    return this.exec({ type: 'insertTable', rows, cols });
  }

  /** Insert an image from a picked file, sized from its intrinsic pixels. */
  async insertImageFile(file: File): Promise<ExecResult> {
    const instance = this.instanceSignal();
    if (!instance) {
      return { ok: false, code: 'unsupported', reason: 'the editor is not mounted yet' };
    }
    const data = new Uint8Array(await file.arrayBuffer());
    const mime = file.type || sniffMime(file.name);
    const size = await imageSize(data, mime);
    // CSS pixels at 96 dpi → points at 72 dpi, capped to a page-friendly width.
    const widthPt = Math.min(((size?.width ?? 480) * 72) / 96, 468);
    const heightPt = size ? (widthPt * size.height) / size.width : 320;
    return this.exec({
      type: 'insertImage',
      data,
      mime: mime as never,
      widthPoints: Math.round(widthPt),
      heightPoints: Math.round(heightPt),
      ...(file.name ? { description: file.name.replace(/\.[^.]+$/, '') } : {}),
    } as unknown as EditorCommand);
  }

  // ── Zoom ─────────────────────────────────────────────────────────────────────────

  /** Set the display scale; leaves any fit mode, as the engine documents. */
  setZoom(zoom: number): ExecResult {
    const instance = this.instanceSignal();
    if (!instance) {
      return { ok: false, code: 'unsupported', reason: 'the editor is not mounted yet' };
    }
    return instance.setZoom(zoom);
  }

  /** Choose between a fixed scale and a viewport-tracking fit. */
  setZoomMode(mode: ZoomMode | 'auto'): ExecResult {
    const instance = this.instanceSignal();
    if (!instance) {
      return { ok: false, code: 'unsupported', reason: 'the editor is not mounted yet' };
    }
    return instance.setZoomMode(mode);
  }

  /** Step to the next rung, or to 100% when already at the top. */
  zoomIn(): ExecResult {
    return this.setZoom(stepZoomLevel(this.zoom(), 'in') ?? 1);
  }

  /** Step to the previous rung. */
  zoomOut(): ExecResult {
    return this.setZoom(stepZoomLevel(this.zoom(), 'out') ?? 0.25);
  }

  /** The zoom a Ctrl/Cmd `0`, `+` or `-` chord asks for, applied when it names one. */
  zoomForShortcut(key: string): boolean {
    const next = zoomLevelForShortcut(key, this.zoom());
    if (next === null) return false;
    this.setZoom(next);
    return true;
  }

  // ── The ref surface ──────────────────────────────────────────────────────────────

  load(document: DocumentSource): void {
    this.instanceSignal()?.load(document);
  }

  async save(): Promise<ArrayBuffer | null> {
    const instance = this.instanceSignal();
    if (!instance) return null;
    return instance.save();
  }

  getDocumentHandle(): ReturnType<DocxEditorRef['getDocumentHandle']> {
    return this.instanceSignal()?.getDocumentHandle() ?? null;
  }

  getEditor(): Editor | null {
    return (this.instanceSignal() as unknown as Editor) ?? null;
  }

  focus(scope?: EditorScope): void {
    this.instanceSignal()?.focus(scope);
  }

  // ── Host actions ─────────────────────────────────────────────────────────────────

  private hostHandler(slot: ChromeSlotId): (() => void | Promise<void>) | null {
    const options = this.optionsSignal();
    switch (slot) {
      case 'file.open':
        return () => this.openFile();
      case 'file.save':
        return () => this.saveFile();
      case 'file.print':
        return () => this.printDocument();
      case 'file.exportMarkdown':
        return options.onExport || options.exportHandlers?.markdown
          ? () => this.exportAs('markdown')
          : null;
      case 'file.exportPdf':
        return options.onExport || options.exportHandlers?.pdf
          ? () => this.exportAs('pdf')
          : null;
      default:
        return null;
    }
  }

  /** Open a document: the host's handler, else a packaged file picker. */
  async openFile(): Promise<void> {
    const options = this.optionsSignal();
    if (options.onOpen) {
      options.onOpen();
      return;
    }
    if (typeof document === 'undefined') return;
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = `.docx,.docm,${DOCX_MIME}`;
    input.style.display = 'none';
    document.body.appendChild(input);
    const chosen = await new Promise<File | null>((resolve) => {
      input.addEventListener('change', () => resolve(input.files?.[0] ?? null), { once: true });
      input.addEventListener('cancel', () => resolve(null), { once: true });
      input.click();
    });
    input.remove();
    if (!chosen) return;
    const bytes = new Uint8Array(await chosen.arrayBuffer());
    this.load(bytes);
  }

  /** Save: the host's handler, else a `.docx` download named after the document. */
  async saveFile(): Promise<void> {
    const options = this.optionsSignal();
    if (options.onSave) {
      options.onSave();
      return;
    }
    try {
      const bytes = await this.save();
      if (!bytes) return;
      download(bytes, `${options.fileName ?? options.title ?? 'document'}.docx`, DOCX_MIME);
    } catch (error) {
      this.notice.set(error instanceof Error ? error.message : String(error));
    }
  }

  /** Export markdown or PDF: the host's handler, else the engine's export lane. */
  async exportAs(format: ChromeExportFormat): Promise<void> {
    const options = this.optionsSignal();
    if (options.onExport) {
      await options.onExport(format);
      return;
    }
    const instance = this.instanceSignal();
    if (!instance) return;
    try {
      const result = await runChromeExport(instance, format, options.exportHandlers ?? {});
      download(
        result.bytes,
        `${options.fileName ?? options.title ?? 'document'}.${result.extension}`,
        result.mimeType
      );
    } catch (error) {
      const t = this.t();
      const message = error instanceof Error ? error.message : String(error);
      this.notice.set(t('toolbar.exportFailed', { message }));
    }
  }

  /** Print: the host's handler, else PDF conversion via the hidden frame, else `window.print`. */
  async printDocument(): Promise<void> {
    const options = this.optionsSignal();
    if (options.onPrint) {
      await options.onPrint();
      return;
    }
    const instance = this.instanceSignal();
    if (instance && options.exportHandlers?.pdf) {
      try {
        const job = await runChromePrint(instance, options.exportHandlers);
        try {
          await job.print();
        } finally {
          job.dispose();
        }
        return;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.notice.set(this.t()('toolbar.printFailedMessage', { message }));
        return;
      }
    }
    if (typeof window !== 'undefined') window.print();
  }

  /** Dismiss the current notice. */
  dismissNotice(): void {
    this.notice.set(null);
  }

  private reportRefusal(reason: string): void {
    const localize = localizeDisabledReason(reason, this.t() as never);
    this.notice.set(localize ?? reason);
  }
}

/**
 * Whether every value in `patch` is already in force.
 *
 * Identity comparison for functions and byte arrays, by design: a host that re-creates a
 * callback every change-detection pass would otherwise re-configure the editor on every pass,
 * which is exactly the loop this guard exists to stop. Hosts keep callback identity stable by
 * binding them once, as the packaged component does.
 */
function shallowEqualOptions(current: DocxEditorOptions, patch: DocxEditorOptions): boolean {
  const keys = Object.keys(patch) as (keyof DocxEditorOptions)[];
  for (const key of keys) {
    if (current[key] !== patch[key]) return false;
  }
  return keys.length > 0;
}

function download(bytes: ArrayBuffer | Uint8Array, fileName: string, mimeType: string): void {
  if (typeof document === 'undefined') return;
  const view = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const part = (view.buffer instanceof ArrayBuffer ? view : view.slice()) as Uint8Array<ArrayBuffer>;
  const url = URL.createObjectURL(new Blob([part], { type: mimeType }));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.rel = 'noopener';
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function sniffMime(name: string): string {
  const extension = name.split('.').pop()?.toLowerCase();
  switch (extension) {
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'webp':
      return 'image/webp';
    case 'svg':
      return 'image/svg+xml';
    case 'bmp':
      return 'image/bmp';
    case 'tif':
    case 'tiff':
      return 'image/tiff';
    default:
      return 'image/png';
  }
}

async function imageSize(
  data: Uint8Array,
  mime: string
): Promise<{ width: number; height: number } | null> {
  if (typeof createImageBitmap !== 'function' || mime === 'image/svg+xml') return null;
  try {
    const bitmap = await createImageBitmap(new Blob([data as Uint8Array<ArrayBuffer>], { type: mime }));
    const size = { width: bitmap.width, height: bitmap.height };
    bitmap.close();
    return size;
  } catch {
    return null;
  }
}
