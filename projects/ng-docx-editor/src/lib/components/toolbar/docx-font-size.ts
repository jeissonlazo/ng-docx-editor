import { ChangeDetectionStrategy, Component, computed, inject } from '@angular/core';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { refusalText } from '../../core/labels';
import { DocxMenuDirective } from '../shared/docx-menu.directive';

/** The sizes Word's picker offers, in points, before the document's own value joins them. */
const SIZE_PRESETS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72];

/**
 * The font-size stepper: `− [value ▾] +`, boxed value that opens the preset list.
 *
 * The engine stores sizes in half-points, so the displayed points are converted on the way
 * out and in — `11.5 pt` is `23`.
 */
@Component({
  selector: 'docx-font-size',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuDirective],
  template: `
    <div class="docx-toolbar__stepper" docxMenu #menu="docxMenu">
      <button
        type="button"
        class="docx-toolbar__stepper-button"
        data-menu-trigger
        [attr.title]="t()('fontSize.decrease')"
        [attr.aria-label]="t()('fontSize.decrease')"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="step(-2)"
      >
        −
      </button>
      <button
        type="button"
        class="docx-toolbar__stepper-value docx-toolbar__stepper-value--boxed"
        data-menu-trigger
        [attr.aria-label]="t()('fontSize.listLabel')"
        [attr.aria-expanded]="menu.open()"
        [attr.title]="title()"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="menu.toggle()"
      >
        {{ display() }}
      </button>
      <button
        type="button"
        class="docx-toolbar__stepper-button"
        data-menu-trigger
        [attr.title]="t()('fontSize.increase')"
        [attr.aria-label]="t()('fontSize.increase')"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="step(2)"
      >
        +
      </button>
      @if (menu.open()) {
        <div class="docx-toolbar__menu docx-toolbar__font-size-menu">
          @for (size of sizes(); track size) {
            <button
              type="button"
              class="docx-toolbar__menu-item"
              data-menu-item
              [attr.aria-label]="size + ' pt'"
              (click)="pick(size)"
            >
              {{ size }}
              @if (size === currentPt()) {
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
      .docx-toolbar__font-size-menu {
        position: absolute;
        top: 100%;
        inset-inline-start: 0;
        z-index: var(--doc-z-overlay);
        margin-top: 4px;
        max-height: 280px;
        overflow-y: auto;
      }
    `,
  ],
})
export class DocxFontSizeComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  readonly slot = 'font.size' as const;
  readonly state = computed(() => this.store.commandState(this.slot));

  readonly currentPt = computed(() => this.store.formatting()?.fontSizePt ?? null);
  readonly display = computed(() => {
    const points = this.currentPt();
    if (points === null) return '';
    return Number.isInteger(points) ? String(points) : points.toFixed(1);
  });
  readonly title = computed(() => {
    const reason = this.state().enabled ? null : refusalText(this.t(), this.state().disabledReason);
    return reason ? `${this.t()('fontSize.listLabel')} — ${reason}` : this.t()('fontSize.listLabel');
  });

  /** The presets, with the document's own size inserted in order when it is not one of them. */
  readonly sizes = computed(() => {
    const current = this.currentPt();
    if (current === null || SIZE_PRESETS.includes(current)) return SIZE_PRESETS;
    return [...SIZE_PRESETS, current].sort((a, b) => a - b);
  });

  pick(points: number): void {
    this.store.setFontSize(Math.round(points * 2));
  }

  step(halfPoints: number): void {
    const current = this.currentPt();
    const next = Math.max(1, (current === null ? 11 : current) * 2 + halfPoints);
    this.store.setFontSize(next);
  }
}
