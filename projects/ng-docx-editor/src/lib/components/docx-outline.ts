import { ChangeDetectionStrategy, Component, computed, inject, input, model } from '@angular/core';
import { DocxEditorStore } from '../core/docx-editor-store';
import { DocxIconComponent } from './shared/docx-icon';

/**
 * The navigation pane's Headings tab: the document's own outline, in document order.
 *
 * Rows come from `Editor.getOutline()` — the engine's headings, not a parsed guess — and a
 * click scrolls the document to that block with `scrollToBlock`. The pane floats over the
 * gutter the stylesheet already reserves, so opening it never re-paginates the document.
 */
@Component({
  selector: 'docx-outline',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxIconComponent],
  template: `
    <div class="docx-nav" [class.docx-nav--open]="open()" [class.docx-nav--overlay]="overlay()">
      @if (!open()) {
        <button
          type="button"
          class="docx-nav__toggle"
          [attr.title]="t()('navigation.openTitle')"
          [attr.aria-label]="t()('navigation.openAriaLabel')"
          [attr.aria-expanded]="false"
          (click)="open.set(true)"
        >
          <docx-icon [paths]="headingsIcon" [size]="18" />
        </button>
      }
      <div class="docx-nav__panel-shell">
        <div class="docx-nav__panel">
          <div class="docx-nav__header">
            <span class="docx-nav__heading">{{ t()('navigation.title') }}</span>
            <button
              type="button"
              class="docx-nav__close"
              [attr.title]="t()('navigation.closeTitle')"
              [attr.aria-label]="t()('navigation.closeAriaLabel')"
              (click)="open.set(false)"
            >
              ✕
            </button>
          </div>
          <div class="docx-nav__list" role="tree">
            @for (heading of outline(); track heading.blockId) {
              <button
                type="button"
                class="docx-nav__heading"
                [class.docx-nav__heading--current]="heading.blockId === activeId()"
                role="treeitem"
                [attr.aria-level]="heading.level"
                [style.padding-inline-start.px]="8 + (heading.level - 1) * 14"
                [attr.title]="heading.text || t()('navigation.headings.noHeadings')"
                (click)="go(heading.blockId)"
              >
                {{ heading.text }}
              </button>
            } @empty {
              <div class="docx-nav__empty">{{ t()('navigation.headings.noHeadings') }}</div>
            }
          </div>
        </div>
      </div>
    </div>
  `,
  styles: [
    `
      :host {
        display: contents;
      }
      .docx-nav {
        --docx-nav-reservation: 328px;
        --docx-nav-inset: 12px;
        --docx-nav-width: 296px;
      }
      .docx-nav__panel-shell {
        border-inline-end: 1px solid var(--doc-border);
        border-radius: 0 8px 8px 0;
        box-shadow: 0 4px 16px var(--doc-shadow);
      }
      .docx-nav__panel {
        display: flex;
        flex-direction: column;
        height: 100%;
        min-height: 0;
      }
      .docx-nav__header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        padding: 10px 12px;
        border-bottom: 1px solid var(--doc-border);
      }
      .docx-nav__heading {
        font-weight: 600;
        font-size: 13px;
      }
      .docx-nav__close {
        border: none;
        background: transparent;
        color: var(--doc-text-muted);
        cursor: pointer;
        font-size: 13px;
        padding: 2px 6px;
        border-radius: 4px;
      }
      .docx-nav__close:hover {
        background: var(--doc-bg-hover);
        color: var(--doc-text);
      }
      .docx-nav__list {
        display: flex;
        flex-direction: column;
        gap: 1px;
        flex: 1;
        min-height: 0;
        overflow-y: auto;
        padding: 6px;
      }
      /* Rows must NOT flex-shrink: a long outline overflows the list, and a shrinkable row in
         a column flex container collapses toward its padding — every heading flattened to a
         sliver. */
      .docx-nav__list > * {
        flex: none;
      }
      button.docx-nav__heading {
        display: block;
        width: 100%;
        padding: 5px 8px;
        border: none;
        border-radius: 4px;
        background: transparent;
        color: var(--doc-text);
        font: inherit;
        font-size: 12.5px;
        font-weight: 400;
        text-align: start;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        cursor: pointer;
      }
      button.docx-nav__heading:hover {
        background: var(--doc-bg-hover);
      }
      .docx-nav__heading--current {
        background: var(--doc-bg-hover);
        font-weight: 600;
      }
      .docx-nav__empty {
        padding: 10px 12px;
        color: var(--doc-text-muted);
        font-size: 12px;
      }
      .docx-nav__toggle {
        position: absolute;
        inset-block-start: 12px;
        inset-inline-start: var(--docx-nav-inset, 12px);
        display: flex;
        align-items: center;
        justify-content: center;
        width: 34px;
        height: 34px;
        padding: 0;
        border: 1px solid var(--doc-border-light);
        border-radius: 50%;
        background: var(--doc-bg-hover);
        color: var(--doc-text-muted);
        cursor: pointer;
        z-index: var(--doc-z-chrome);
      }
      .docx-nav__toggle:hover {
        color: var(--doc-text);
      }
    `,
  ],
})
export class DocxOutlineComponent {
  private readonly store = inject(DocxEditorStore);

  /** The pane's open state, two-way bindable. */
  readonly open = model(false);
  /**
   * Whether the pane covers the page instead of sitting in the gutter.
   *
   * Default `false`: the pane floats over the gutter the stylesheet already reserves beside
   * the centred page, so opening it never re-paginates the document. A host on a very narrow
   * window sets `true` and the panel spans the viewport instead.
   */
  readonly overlay = input(false);

  readonly t = this.store.t;
  readonly outline = computed(() => this.store.outline());
  readonly activeId = model<string | null>(null);

  readonly headingsIcon = [
    'M120-240v-80h240v80H120Zm0-200v-80h480v80H120Zm0-200v-80h720v80H120Z',
  ];

  go(blockId: string): void {
    this.activeId.set(blockId);
    this.store.instance()?.scrollToBlock(blockId, { block: 'start' });
  }
}
