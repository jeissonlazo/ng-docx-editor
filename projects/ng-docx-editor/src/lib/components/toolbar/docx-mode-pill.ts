import { ChangeDetectionStrategy, Component, computed, inject } from '@angular/core';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { refusalText } from '../../core/labels';
import { DocxMenuDirective } from '../shared/docx-menu.directive';

type PillMode = 'editing' | 'suggesting' | 'viewing';

const MODES: readonly PillMode[] = ['editing', 'suggesting', 'viewing'];

/**
 * The editing-mode pill: Editing · Suggesting · Viewing.
 *
 * Enabled state and the current value come from the engine (`review.editingMode`), and each
 * row carries its own refusal — suggesting needs a review module and an author, so a document
 * that cannot enter it says so instead of silently staying put.
 */
@Component({
  selector: 'docx-mode-pill',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuDirective],
  template: `
    <div class="docx-toolbar__mode" docxMenu #menu="docxMenu">
      <button
        type="button"
        class="docx-toolbar__picker docx-toolbar__mode-trigger"
        data-menu-trigger
        [attr.aria-label]="t()('editingMode.label')"
        [attr.aria-expanded]="menu.open()"
        [attr.title]="title()"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="menu.toggle()"
      >
        <span class="docx-toolbar__mode-text">{{ modeLabel(current()) }}</span>
        <span class="docx-toolbar__picker-caret">▾</span>
      </button>
      @if (menu.open()) {
        <div class="docx-toolbar__menu docx-toolbar__mode-menu">
          @for (mode of modes; track mode) {
            <button
              type="button"
              class="docx-toolbar__menu-item docx-toolbar__mode-item"
              data-menu-item
              [attr.aria-disabled]="mode === current() ? 'true' : null"
              (click)="pick(mode)"
            >
              <span class="docx-toolbar__mode-label">{{ modeLabel(mode) }}</span>
              <span class="docx-toolbar__mode-hint">{{ modeHint(mode) }}</span>
              @if (mode === current()) {
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
      .docx-toolbar__mode-menu {
        position: absolute;
        top: 100%;
        inset-inline-start: 0;
        z-index: var(--doc-z-overlay);
        margin-top: 4px;
        min-width: 200px;
      }
      .docx-toolbar__mode-item {
        display: grid;
        grid-template-columns: 1fr auto;
        column-gap: 8px;
      }
      .docx-toolbar__mode-hint {
        grid-column: 1 / -1;
        color: var(--doc-text-muted);
        font-size: 11px;
      }
    `,
  ],
})
export class DocxModePillComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  readonly modes = MODES;
  readonly slot = 'review.editingMode' as const;
  readonly state = computed(() => this.store.commandState(this.slot));

  readonly current = computed<PillMode>(() => {
    const value = this.state().value;
    if (value === 'editing' || value === 'suggesting' || value === 'viewing') return value;
    return (this.store.editingMode() as PillMode) ?? 'editing';
  });

  readonly title = computed(() => {
    const reason = this.state().enabled ? null : refusalText(this.t(), this.state().disabledReason);
    const label = this.t()('editingMode.label');
    return reason ? `${label} — ${reason}` : label;
  });

  modeLabel(mode: PillMode): string {
    return this.t()(`editingMode.${mode}`);
  }

  modeHint(mode: PillMode): string {
    return this.t()(`editingMode.${mode}Hint`);
  }

  pick(mode: PillMode): void {
    if (mode === this.current()) return;
    this.store.setEditingMode(mode);
  }
}
