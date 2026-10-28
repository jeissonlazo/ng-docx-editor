import {
  ChangeDetectionStrategy,
  Component,
  computed,
  effect,
  inject,
  input,
  output,
  signal,
  untracked,
} from '@angular/core';
import type { Editor, EditorFontError, DocumentChange } from '@docx-editor.dev/core/contracts/editor';
import type { ChromeExportFormat } from '@docx-editor.dev/core/editor';
import type { Translations } from '@docx-editor.dev/i18n';
import { DocxEditorStore } from '../core/docx-editor-store';
import { DOCX_EDITOR_SHORTCUTS } from '../core/tokens';
import type { DocxEditorOptions, EditorMode } from '../core/types';
import type { DocxFonts } from '../core/types';
import { DocxEditorContentComponent } from './docx-editor-content';
import { DocxEditorViewportComponent } from './docx-editor-viewport';
import { DocxToolbarComponent } from './toolbar/docx-toolbar';
import { DocxMenuBarComponent } from './menu/docx-menu-bar';
import { DocxOutlineComponent } from './docx-outline';
import { DocxPageIndicatorComponent } from './docx-page-indicator';

/**
 * The batteries-included Angular editor: chrome, English labels, and a painted editable
 * document, with no further configuration.
 *
 * `<docx-editor [document]="bytes" />` is complete. Everything below is about what you can
 * change, not what you must supply. The chrome is composed from the same primitives a host
 * would use (`provideDocxEditor` + `<docx-editor-toolbar>` + `<docx-editor-viewport>` +
 * `<docx-editor-content>`), so a host that outgrows it drops to those with no behaviour change.
 *
 * @example
 * ```html
 * <docx-editor [document]="bytes" title="Contract.docx" [onSave]="save" />
 * ```
 */
@Component({
  selector: 'docx-editor',
  changeDetection: ChangeDetectionStrategy.OnPush,
  exportAs: 'docxEditor',
  providers: [DocxEditorStore],
  imports: [
    DocxEditorContentComponent,
    DocxEditorViewportComponent,
    DocxToolbarComponent,
    DocxMenuBarComponent,
    DocxOutlineComponent,
    DocxPageIndicatorComponent,
  ],
  template: `
    <div
      class="docx-editor ngdx-shell"
      [class.dark]="dark()"
      [attr.dir]="dir()"
      (keydown)="onKeyDown($event)"
    >
      @if (chrome()) {
        <div class="ngdx-chrome-band">
          <div class="ngdx-title-bar">
            <div class="ngdx-title-block">
              @if (onTitleChange()) {
                <input
                  class="ngdx-title-input"
                  [attr.aria-label]="t()('titleBar.documentNameAriaLabel')"
                  [attr.placeholder]="t()('titleBar.untitled')"
                  [value]="title() ?? ''"
                  (input)="titleChanged($any($event.target).value)"
                />
              } @else {
                <span class="ngdx-title-text">{{ title() ?? t()('titleBar.untitled') }}</span>
              }
              @if (menu()) {
                <docx-menu-bar />
              }
            </div>
          </div>
          <div class="ngdx-toolbar-row">
            <docx-toolbar />
          </div>
        </div>
      }

      @if (notice(); as message) {
        <div class="ngdx-notice" role="status">
          <span>{{ message }}</span>
          <button
            type="button"
            class="ngdx-notice-close"
            [attr.aria-label]="t()('common.dismiss')"
            (click)="store.dismissNotice()"
          >
            ✕
          </button>
        </div>
      }

      <div class="ngdx-workspace">
        @if (navigation()) {
          <docx-outline [(open)]="outlineOpen" />
        }
        <docx-editor-viewport>
          <docx-editor-content />
        </docx-editor-viewport>
        <docx-page-indicator />
        @if (loading()) {
          <div class="docx-editor__loading ngdx-loading" role="status">
            <div class="ngdx-loading-card">
              <div class="docx-editor__loading-spinner ngdx-spinner"></div>
              <span>{{ t()('loading.label') }}</span>
            </div>
          </div>
        }
        @if (openFailure(); as failure) {
          @if (!failureDismissed()) {
            <div class="docx-dialog-overlay ngdx-dialog-overlay">
              <div class="docx-dialog" role="dialog" aria-modal="true">
                <div class="docx-dialog__header">
                  <span class="docx-dialog__title">{{ t()('openFailure.title') }}</span>
                </div>
                <div class="docx-dialog__body">
                  <p class="docx-dialog__message">{{ failureMessage() }}</p>
                  <p class="docx-dialog__code">
                    {{ t()('openFailure.code', { code: failure.code }) }}
                  </p>
                </div>
                <div class="docx-dialog__footer">
                  <button
                    type="button"
                    class="docx-dialog__button docx-dialog__button--primary"
                    (click)="failureDismissed.set(true)"
                  >
                    {{ t()('common.close') }}
                  </button>
                </div>
              </div>
            </div>
          }
        }
      </div>
    </div>
  `,
  styles: [
    `
      :host {
        display: block;
        height: 100%;
        min-height: 0;
      }
      .ngdx-shell {
        display: flex;
        flex-direction: column;
        height: 100%;
        width: 100%;
        min-height: 0;
        background: var(--doc-bg);
        color: var(--doc-text);
      }
      .ngdx-chrome-band {
        flex: none;
        display: flex;
        flex-direction: column;
        position: relative;
        z-index: var(--doc-z-overlay);
        background: var(--doc-surface);
        border-bottom: 1px solid var(--doc-border);
        box-shadow: 0 1px 3px var(--doc-shadow-subtle);
      }
      .ngdx-title-bar {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 8px 12px 0;
      }
      .ngdx-title-block {
        display: flex;
        flex-direction: column;
        flex: 1;
        min-width: 0;
        gap: 2px;
      }
      .ngdx-title-text,
      .ngdx-title-input {
        min-width: 0;
        padding: 2px 6px;
        font: inherit;
        font-size: 14px;
        color: var(--doc-text);
      }
      .ngdx-title-input {
        border: 1px solid transparent;
        border-radius: 4px;
        background: transparent;
      }
      .ngdx-title-input:focus {
        border-color: var(--doc-border);
        outline: none;
      }
      .ngdx-toolbar-row {
        padding: 8px 12px;
      }
      .ngdx-notice {
        flex: none;
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 6px 12px;
        font-size: 12px;
        color: var(--doc-text);
        background: var(--doc-bg-subtle);
        border-bottom: 1px solid var(--doc-border);
      }
      .ngdx-notice span {
        flex: 1;
      }
      .ngdx-notice-close {
        border: none;
        background: transparent;
        color: inherit;
        cursor: pointer;
      }
      .ngdx-workspace {
        position: relative;
        display: flex;
        flex: 1;
        min-height: 0;
        min-width: 0;
      }
      .ngdx-loading {
        position: absolute;
        inset: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        background: var(--doc-bg);
        z-index: var(--doc-z-overlay);
      }
      .ngdx-loading-card {
        display: flex;
        align-items: center;
        gap: 10px;
        color: var(--doc-text-muted);
        font-size: 13px;
      }
      .ngdx-spinner {
        width: 18px;
        height: 18px;
        border: 2px solid var(--doc-border);
        border-top-color: var(--doc-accent);
        border-radius: 50%;
        animation: ngdx-spin 0.9s linear infinite;
      }
      @keyframes ngdx-spin {
        to {
          transform: rotate(360deg);
        }
      }
      .ngdx-dialog-overlay {
        position: absolute;
        inset: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        background: rgba(0, 0, 0, 0.35);
        z-index: calc(var(--doc-z-overlay) + 10);
      }
    `,
  ],
})
export class DocxEditorComponent {
  /** The editor's state container; also the imperative surface for `#ref="docxEditor"`. */
  readonly store = inject(DocxEditorStore);
  private readonly shortcuts = inject(DOCX_EDITOR_SHORTCUTS, { optional: true }) ?? true;

  // ── Document source ───────────────────────────────────────────────────────────────
  /** A document to load: DOCX bytes, `'blank'`, or an existing handle. */
  readonly document = input<DocxEditorOptions['document']>(undefined);
  /** Font sources for Word-accurate shaping and pagination. */
  readonly fonts = input<DocxFonts | undefined>(undefined);
  /** Explicit shaping byte ceiling. */
  readonly fontShaping = input<DocxEditorOptions['fontShaping']>(undefined);
  /** Author for comments, replies and tracked changes. */
  readonly author = input<string | undefined>(undefined);
  /** BCP-47 locale for regional date input and generated labels. */
  readonly locale = input<string | undefined>(undefined);
  /** Capability modules (pro review module, custom nodes). Sampled at mount. */
  readonly modules = input<DocxEditorOptions['modules']>(undefined);
  /** The host editing mode. Omitted, the document's `w:trackRevisions` decides. */
  readonly mode = input<EditorMode | undefined>(undefined);
  /** A fixed display scale. */
  readonly zoom = input<number | undefined>(undefined);
  /** Where the display scale comes from. */
  readonly zoomMode = input<DocxEditorOptions['zoomMode']>(undefined);
  /** Controlled viewer-markup settings. */
  readonly revisionMarkup = input<DocxEditorOptions['revisionMarkup']>(undefined);
  /** Initial revision projection. */
  readonly reviewDisplayMode = input<DocxEditorOptions['reviewDisplayMode']>(undefined);
  /** Optional image decode port. */
  readonly imageDecodePort = input<DocxEditorOptions['imageDecodePort']>(undefined);

  // ── Chrome ────────────────────────────────────────────────────────────────────────
  /** Renders the packaged chrome around the document. Default `true`. */
  readonly chrome = input(true);
  /** Renders the menu bar. Default `true`. */
  readonly menu = input(true);
  /** Renders the headings pane. Default `true`. */
  readonly navigation = input(true);
  /** Chrome colour mode; `'system'` follows the OS. */
  readonly colorMode = input<'light' | 'dark' | 'system'>('light');
  /** Catalogue for this editor, merged onto the inherited one. */
  readonly i18n = input<Translations | undefined>(undefined);
  /** Forces the chrome writing direction. */
  readonly dir = input<'ltr' | 'rtl' | undefined>(undefined);
  /** Document name shown in the title bar. */
  readonly title = input<string | undefined>(undefined);

  // ── Host actions (callbacks, like React's props) ─────────────────────────────────
  /** Open handler for File › Open. Without it a packaged file picker loads the bytes. */
  readonly onOpen = input<(() => void) | undefined>(undefined);
  /** Save handler for File › Save. Without it the bytes download under `title`. */
  readonly onSave = input<(() => void) | undefined>(undefined);
  /** Export handler for File › Export. Without it those rows are unavailable. */
  readonly onExport = input<((format: ChromeExportFormat) => void | Promise<void>) | undefined>(
    undefined
  );
  /** Print handler. Without it the packaged print runs `window.print()`. */
  readonly onPrint = input<(() => void | Promise<void>) | undefined>(undefined);
  /** Base file name used by the packaged downloads. */
  readonly fileName = input<string | undefined>(undefined);
  /** Title editor; without it the title is read-only. */
  readonly onTitleChange = input<((title: string) => void) | undefined>(undefined);

  // ── Events ────────────────────────────────────────────────────────────────────────
  /** Fired once per instance, after the first document paints. */
  readonly ready = output<Editor>();
  /** Fired when the document changes. */
  readonly change = output<DocumentChange>();
  /** Fired with the typed font failure the engine rejected. */
  readonly fontError = output<EditorFontError>();
  /** Fired when the title editor commits a new name. */
  readonly titleChange = output<string>();

  /** The headings pane's open state. */
  readonly outlineOpen = signal(false);
  /** Whether the open-failure dialog has been dismissed by the reader. */
  readonly failureDismissed = signal(false);

  readonly t = this.store.t;
  readonly dark = this.store.isDark;
  readonly loading = this.store.isLoading;
  readonly notice = this.store.notice;
  readonly openFailure = this.store.openFailure;

  readonly failureMessage = computed(() => {
    const failure = this.openFailure();
    const key = failure ? `openFailure.reasons.${failure.reason}` : 'openFailure.reasons.unknown';
    return this.t()(key);
  });

  constructor() {
    // One effect for every input. The values are READ here (that is what makes the effect
    // re-run when an input moves) and then handed to the store inside `untracked`, so the
    // store's own signals — which it necessarily reads while applying them — never become
    // dependencies of this effect and can never re-trigger it.
    effect(() => {
      const patch: DocxEditorOptions = {
        document: this.document(),
        fonts: this.fonts(),
        fontShaping: this.fontShaping(),
        imageDecodePort: this.imageDecodePort(),
        author: this.author(),
        locale: this.locale(),
        modules: this.modules(),
        mode: this.mode(),
        zoom: this.zoom(),
        zoomMode: this.zoomMode(),
        revisionMarkup: this.revisionMarkup(),
        reviewDisplayMode: this.reviewDisplayMode(),
        i18n: this.i18n(),
        dir: this.dir(),
        title: this.title(),
        colorMode: this.colorMode(),
        fileName: this.fileName(),
        onOpen: this.onOpen(),
        onSave: this.onSave(),
        onExport: this.onExport(),
        onPrint: this.onPrint(),
        onReady: this.emitReady,
        onChange: this.emitChange,
        onFontError: this.emitFontError,
      };
      untracked(() => this.store.configure(patch));
    });
  }

  // Stable callback identities: the store compares option values by identity to stay
  // idempotent, and a closure re-created per effect run would defeat that.
  private readonly emitReady = (editor: Editor) => this.ready.emit(editor);
  private readonly emitChange = (change: DocumentChange) => this.change.emit(change);
  private readonly emitFontError = (error: EditorFontError) => this.fontError.emit(error);

  // ── Imperative surface (`#ref="docxEditor"`) ─────────────────────────────────────

  /** Load a document. */
  load(document: NonNullable<DocxEditorOptions['document']>): void {
    this.store.load(document);
  }

  /** Serialize the current document. */
  save(): Promise<ArrayBuffer | null> {
    return this.store.save();
  }

  /** Identity and revision of the loaded document. */
  getDocumentHandle(): ReturnType<DocxEditorStore['getDocumentHandle']> {
    return this.store.getDocumentHandle();
  }

  /** The full `Editor` facade. */
  getEditor(): Editor | null {
    return this.store.getEditor();
  }

  /** Focus the pages. */
  focus(): void {
    this.store.focus();
  }

  /** Run a typed command. */
  exec(...args: Parameters<DocxEditorStore['exec']>): ReturnType<DocxEditorStore['exec']> {
    return this.store.exec(...args);
  }

  /** The current read model. */
  snapshot(...args: Parameters<DocxEditorStore['snapshot']>) {
    return this.store.snapshot(...args);
  }

  /** Open the file picker (or the host's handler). */
  openFile(): Promise<void> {
    return this.store.openFile();
  }

  /** Save through the host's handler, or download the bytes. */
  saveFile(): Promise<void> {
    return this.store.saveFile();
  }

  titleChanged(value: string): void {
    this.onTitleChange()?.(value);
    this.titleChange.emit(value);
  }

  onKeyDown(event: KeyboardEvent): void {
    if (!this.shortcuts) return;
    if (!(event.metaKey || event.ctrlKey) || event.altKey) return;
    const target = event.target;
    if (target instanceof Element && target.closest('input, textarea, select')) return;
    switch (event.key.toLowerCase()) {
      case 's':
        event.preventDefault();
        void this.store.saveFile();
        break;
      case 'o':
        event.preventDefault();
        void this.store.openFile();
        break;
      case 'p':
        event.preventDefault();
        void this.store.printDocument();
        break;
      default:
        break;
    }
  }
}
