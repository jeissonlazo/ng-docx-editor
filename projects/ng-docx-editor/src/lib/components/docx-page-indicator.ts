import { ChangeDetectionStrategy, Component, computed, inject } from '@angular/core';
import { DocxEditorStore } from '../core/docx-editor-store';

/**
 * The floating `{current} of {total}` readout beside the scrollbar, shown for a document with
 * more than one page.
 *
 * The engine publishes the current page itself (`snapshot().page`), so this is a read of state
 * the pages already own rather than a scroll-position guess.
 */
@Component({
  selector: 'docx-page-indicator',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div
      class="docx-editor__page-number"
      role="status"
      aria-live="polite"
      [attr.data-visible]="visible() ? 'true' : 'false'"
    >
      {{ text() }}
    </div>
  `,
  styles: [
    `
      :host {
        position: absolute;
        inset-inline-end: 0;
        inset-block: 0;
        pointer-events: none;
      }
      .docx-editor__page-number {
        background-color: var(--doc-overlay);
        color: #fff;
        padding: 6px 12px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 500;
        white-space: nowrap;
        user-select: none;
        transition: opacity 0.3s ease;
      }
    `,
  ],
})
export class DocxPageIndicatorComponent {
  private readonly store = inject(DocxEditorStore);

  readonly visible = computed(() => {
    const page = this.store.page();
    return page.total > 1 && page.current > 0;
  });
  readonly text = computed(() =>
    this.store.t()('viewer.pageIndicator', {
      current: this.store.page().current,
      total: this.store.page().total,
    })
  );
}
