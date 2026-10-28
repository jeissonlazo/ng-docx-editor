import { ChangeDetectionStrategy, Component, computed, inject } from '@angular/core';
import { DocxEditorStore } from '../core/docx-editor-store';

/**
 * The scroll container around the painted pages.
 *
 * The class list is LOAD-BEARING, not styling sugar:
 * - `docx-editor-one-surface docx-editor-one-surface__viewport` carry the one-surface chrome
 *   geometry from the engine's stylesheet;
 * - `docx-editor__scroll-container` is how the engine FINDS its scroller — the paginated
 *   pages locate the nearest ancestor with that class for scroll rematerialization and page
 *   visibility. Without it the engine falls back to document scrolling and virtualization
 *   degrades.
 *
 * Zoom shortcuts are claimed in the CAPTURE phase, like React's viewport: the engine's keymap
 * binds the same Ctrl/Cmd chords to character formatting inside the pages, and zoom is the
 * one that must win on those keys.
 */
@Component({
  selector: 'docx-editor-viewport',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div
      class="docx-editor-one-surface docx-editor-one-surface__viewport docx-editor__scroll-container"
      [class.ngdx-viewport]="true"
      tabindex="0"
      role="region"
      [attr.aria-label]="t()('editor.documentViewport')"
      [attr.data-zoom-fit]="fit() ? '' : null"
      (keydown)="onKeyDown($event)"
    >
      <ng-content />
    </div>
  `,
  styles: [
    `
      :host {
        display: flex;
        flex: 1;
        min-height: 0;
        min-width: 0;
        position: relative;
      }
      .docx-editor__scroll-container {
        flex: 1;
        min-height: 0;
        min-width: 0;
        overflow: auto;
        overflow-anchor: none;
      }
    `,
  ],
})
export class DocxEditorViewportComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  readonly fit = computed(() => this.store.zoomMode()?.type === 'fit');

  onKeyDown(event: KeyboardEvent): void {
    if (!(event.metaKey || event.ctrlKey) || event.altKey) return;
    if (!this.store.instance()) return;
    const target = event.target;
    if (target instanceof Element && target.closest('input, textarea, select, [role="dialog"]')) {
      return;
    }
    if (this.store.zoomForShortcut(event.key)) event.preventDefault();
  }
}
