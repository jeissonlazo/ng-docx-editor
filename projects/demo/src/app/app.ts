import { ChangeDetectionStrategy, Component, computed, signal, viewChild } from '@angular/core';
import { defaultFonts } from '@docx-editor.dev/fonts';
import { de, en, es, fr, ptBR } from '@docx-editor.dev/i18n';
import type { Translations } from '@docx-editor.dev/i18n';
import type { DocumentSource } from 'ng-docx-editor';
import { DocxEditorComponent } from 'ng-docx-editor';
import type { DocxFonts } from 'ng-docx-editor';
import { ComposedEditorComponent } from './composed-editor';

type LocaleId = 'en' | 'es' | 'pt-BR' | 'de' | 'fr';

const LOCALES: Record<LocaleId, Translations> = {
  en,
  es,
  'pt-BR': ptBR,
  de,
  fr,
};

const LOCALE_LABELS: Record<LocaleId, string> = {
  en: 'English',
  es: 'Español',
  'pt-BR': 'Português (BR)',
  de: 'Deutsch',
  fr: 'Français',
};

/**
 * The demo host: an app bar that exercises the public API, and the editor itself in either
 * composition mode — the packaged `<docx-editor>` or the provider-first primitives, which
 * render the identical chrome from the same store.
 */
@Component({
  selector: 'app-root',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxEditorComponent, ComposedEditorComponent],
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App {
  private readonly editor = viewChild(DocxEditorComponent);

  readonly localeIds = Object.keys(LOCALE_LABELS) as LocaleId[];
  readonly localeLabels = LOCALE_LABELS;

  readonly locale = signal<LocaleId>('en');
  readonly colorMode = signal<'light' | 'dark' | 'system'>('light');
  readonly composed = signal(false);
  readonly fonts = signal<DocxFonts | undefined>(undefined);
  readonly document = signal<DocumentSource | undefined>(undefined);
  readonly title = signal('sample.docx');
  readonly pages = signal(0);
  readonly revision = signal(0);
  readonly dirty = signal(false);
  readonly busy = signal(false);

  readonly i18n = computed(() => LOCALES[this.locale()]);

  constructor() {
    void this.loadFonts();
    void this.loadSample();
  }

  private async loadFonts(): Promise<void> {
    try {
      this.fonts.set(await defaultFonts());
    } catch {
      // Without fonts the engine still paints; wrap points are estimated instead.
    }
  }

  async loadSample(): Promise<void> {
    this.busy.set(true);
    try {
      const response = await fetch('sample.docx');
      const bytes = new Uint8Array(await response.arrayBuffer());
      this.document.set(bytes);
      this.title.set('sample.docx');
      this.dirty.set(false);
    } finally {
      this.busy.set(false);
    }
  }

  async open(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (!file) return;
    this.document.set(new Uint8Array(await file.arrayBuffer()));
    this.title.set(file.name);
    this.dirty.set(false);
  }

  newDocument(): void {
    this.document.set('blank');
    this.title.set('Untitled.docx');
    this.dirty.set(false);
  }

  async save(): Promise<void> {
    const bytes = await this.editor()?.save();
    if (!bytes) return;
    const url = URL.createObjectURL(
      new Blob([bytes], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      })
    );
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = this.title();
    anchor.click();
    URL.revokeObjectURL(url);
    this.dirty.set(false);
  }

  undo(): void {
    this.editor()?.exec({ type: 'undo' });
  }

  redo(): void {
    this.editor()?.exec({ type: 'redo' });
  }

  onChange(change: { revision: number; source?: string }): void {
    this.revision.set(change.revision);
    if (change.source !== 'load') this.dirty.set(true);
    const page = this.editor()?.snapshot().page;
    if (page) this.pages.set(page.total);
  }

  onReady(): void {
    const page = this.editor()?.snapshot().page;
    if (page) this.pages.set(page.total);
  }

  cycleColorMode(): void {
    const order = ['light', 'dark', 'system'] as const;
    const next = order[(order.indexOf(this.colorMode()) + 1) % order.length];
    this.colorMode.set(next);
  }
}
