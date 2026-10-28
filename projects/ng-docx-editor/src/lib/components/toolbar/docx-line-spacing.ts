import { ChangeDetectionStrategy, Component, computed, inject } from '@angular/core';
import { lineSpacingMenuRowLabel, lineSpacingMenuRows } from '@docx-editor.dev/core/editor';
import type { LineSpacingMenuRow } from '@docx-editor.dev/core/editor';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { refusalText, slotLabel } from '../../core/labels';
import { DocxMenuDirective } from '../shared/docx-menu.directive';

/**
 * The line-spacing picker: the presets plus, ticked, whatever the selection already has.
 *
 * The rows come from the engine's own model (`lineSpacingMenuRows`), so an authored
 * `Exactly 12 pt` shows as the current row instead of a menu that ticks nothing.
 */
@Component({
  selector: 'docx-line-spacing',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuDirective],
  template: `
    <div class="docx-toolbar__line-spacing" docxMenu #menu="docxMenu">
      <button
        type="button"
        class="docx-toolbar__picker docx-toolbar__line-spacing-trigger"
        data-menu-trigger
        [attr.aria-label]="label()"
        [attr.aria-expanded]="menu.open()"
        [attr.title]="title()"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="menu.toggle()"
      >
        {{ t()('lineSpacing.label') }}
        <span class="docx-toolbar__picker-caret">▾</span>
      </button>
      @if (menu.open()) {
        <div class="docx-toolbar__menu docx-toolbar__line-spacing-menu">
          @for (row of rows(); track row.text + row.spacing.rule) {
            <button
              type="button"
              class="docx-toolbar__menu-item"
              data-menu-item
              [attr.aria-disabled]="row.command === null ? 'true' : null"
              (click)="pick(row)"
            >
              {{ rowLabel(row) }}
              @if (row.selected) {
                <span class="docx-toolbar__menu-check">✓</span>
              }
            </button>
          }
        </div>
      }
    </div>
  `,
  styles: [
    `
      :host {
        position: relative;
        display: inline-flex;
      }
      .docx-toolbar__line-spacing-menu {
        position: absolute;
        top: 100%;
        inset-inline-start: 0;
        z-index: var(--doc-z-overlay);
        margin-top: 4px;
        min-width: 160px;
      }
    `,
  ],
})
export class DocxLineSpacingComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  readonly slot = 'list.lineSpacing' as const;
  readonly state = computed(() => this.store.commandState(this.slot));
  readonly label = computed(() => slotLabel(this.store.t(), this.slot));
  readonly title = computed(() => {
    const reason = this.state().enabled ? null : refusalText(this.t(), this.state().disabledReason);
    return reason ? `${this.label()} — ${reason}` : this.label();
  });

  private readonly current = computed(() => this.store.formatting()?.lineSpacing ?? null);
  readonly rows = computed(() => lineSpacingMenuRows(this.current()));

  rowLabel(row: LineSpacingMenuRow): string {
    return lineSpacingMenuRowLabel(row, (key) => this.t()(key));
  }

  pick(row: LineSpacingMenuRow): void {
    if (row.command) this.store.exec(row.command);
  }
}
