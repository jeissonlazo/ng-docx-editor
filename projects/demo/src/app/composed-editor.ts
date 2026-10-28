import { ChangeDetectionStrategy, Component, effect, inject, input, model, untracked } from '@angular/core';
import type { DocumentSource } from 'ng-docx-editor';
import type { DocxFonts } from 'ng-docx-editor';
import {
  DocxEditorContentComponent,
  DocxEditorStore,
  DocxEditorViewportComponent,
  DocxMenuBarComponent,
  DocxOutlineComponent,
  DocxToolbarComponent,
  provideDocxEditor,
} from 'ng-docx-editor';

/**
 * The provider-first composition: `provideDocxEditor()` owns the instance at this component's
 * injector level, and every chrome part below reads it through DI.
 *
 * This is the Angular equivalent of React's `<DocxEditor.Root>` + parts, and it demonstrates
 * that the packaged `<docx-editor>` is sugar over exactly these primitives.
 */
@Component({
  selector: 'demo-composed-editor',
  changeDetection: ChangeDetectionStrategy.OnPush,
  providers: [provideDocxEditor()],
  imports: [
    DocxToolbarComponent,
    DocxMenuBarComponent,
    DocxOutlineComponent,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
  ],
  template: `
    <div class="docx-editor composed-shell">
      <div class="composed-chrome">
        <docx-menu-bar />
        <docx-toolbar />
      </div>
      <div class="composed-body">
        <docx-outline [(open)]="outlineOpen" />
        <docx-editor-viewport>
          <docx-editor-content />
        </docx-editor-viewport>
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
      .composed-shell {
        display: flex;
        flex-direction: column;
        height: 100%;
        background: var(--doc-bg);
      }
      .composed-chrome {
        flex: none;
        padding: 8px 12px;
        background: var(--doc-surface);
        border-bottom: 1px solid var(--doc-border);
      }
      .composed-body {
        position: relative;
        display: flex;
        flex: 1;
        min-height: 0;
      }
      .composed-body docx-editor-viewport {
        flex: 1;
        min-width: 0;
      }
    `,
  ],
})
export class ComposedEditorComponent {
  private readonly store = inject(DocxEditorStore);

  readonly document = input<DocumentSource | undefined>(undefined);
  readonly fonts = input<DocxFonts | undefined>(undefined);
  readonly outlineOpen = model(false);

  constructor() {
    // The same options the packaged component forwards, applied through the public store.
    effect(() => {
      const document = this.document();
      const fonts = this.fonts();
      untracked(() => this.store.configure({ document, fonts }));
    });
  }
}
