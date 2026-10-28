import { ChangeDetectionStrategy, Component, computed, inject } from '@angular/core';
import type { ZoomMode } from '@docx-editor.dev/core/contracts/editor';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { stepZoomLevel, zoomMenuRows, type ZoomMenuRow } from '../../core/zoom';
import { DocxMenuDirective } from '../shared/docx-menu.directive';

/**
 * The zoom stepper: `− 100% ▾ +`, the middle opening the packaged level list.
 *
 * Zoom is engine state, not a host transform, so the same number drives the painted pages,
 * the hit testing and this readout — the control cannot drift from the page.
 */
@Component({
  selector: 'docx-zoom-stepper',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuDirective],
  template: `
    <div class="docx-toolbar__stepper docx-toolbar__zoom" docxMenu #menu="docxMenu">
      <button
        type="button"
        class="docx-toolbar__stepper-button"
        data-menu-trigger
        [attr.title]="t()('zoom.zoomOut')"
        [attr.aria-label]="t()('zoom.zoomOut')"
        [disabled]="!canZoomOut()"
        (mousedown)="$event.preventDefault()"
        (click)="stepOut()"
      >
        −
      </button>
      <button
        type="button"
        class="docx-toolbar__stepper-value"
        data-menu-trigger
        [attr.aria-label]="ariaLabel()"
        [attr.aria-expanded]="menu.open()"
        [attr.title]="ariaLabel()"
        (mousedown)="$event.preventDefault()"
        (click)="menu.toggle()"
      >
        {{ percent() }}%
      </button>
      <button
        type="button"
        class="docx-toolbar__stepper-button"
        data-menu-trigger
        [attr.title]="t()('zoom.zoomIn')"
        [attr.aria-label]="t()('zoom.zoomIn')"
        [disabled]="!canZoomIn()"
        (mousedown)="$event.preventDefault()"
        (click)="stepIn()"
      >
        +
      </button>
      @if (menu.open()) {
        <div class="docx-toolbar__menu docx-toolbar__zoom-menu">
          @for (row of rows(); track row.id) {
            <button
              type="button"
              class="docx-toolbar__menu-item"
              data-menu-item
              (click)="pick(row)"
            >
              {{ row.label }}
              @if (selected(row)) {
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
      /* The middle cell is a BUTTON (it opens the level menu), so it needs the bar's own
         chrome spelled out: without this it keeps the UA's button face and, in the dark
         chrome, paints a light slab with light text on it. */
      .docx-toolbar__stepper-value {
        border: none;
        background: transparent;
        color: var(--doc-text);
        font: inherit;
        cursor: pointer;
        border-radius: 4px;
      }
      .docx-toolbar__stepper-value:hover {
        background: var(--doc-border);
      }
      .docx-toolbar__zoom-menu {
        position: absolute;
        top: 100%;
        inset-inline-start: 50%;
        transform: translateX(-50%);
        z-index: var(--doc-z-overlay);
        margin-top: 4px;
        min-width: 150px;
      }
    `,
  ],
})
export class DocxZoomStepperComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  readonly percent = computed(() => Math.round(this.store.zoom() * 100));
  readonly ariaLabel = computed(() => this.t()('zoom.ariaLabel', { label: `${this.percent()}%` }));
  readonly canZoomIn = computed(() => this.percent() < 500);
  readonly canZoomOut = computed(() => this.percent() > 10);
  readonly rows = computed(() => zoomMenuRows((key) => this.t()(key)));

  stepIn(): void {
    this.store.zoomIn();
  }

  stepOut(): void {
    this.store.zoomOut();
  }

  pick(row: ZoomMenuRow): void {
    if (row.zoom !== undefined) this.store.setZoom(row.zoom);
    else if (row.mode) this.store.setZoomMode(row.mode as ZoomMode);
  }

  /** Whether the row is the state in force: the level, or the mode. */
  selected(row: ZoomMenuRow): boolean {
    const mode = this.store.zoomMode();
    if (row.mode) {
      const current = mode ?? { type: 'fixed' as const };
      const target = row.mode;
      return (
        current.type === target.type &&
        (current.type !== 'fit' ||
          target.type !== 'fit' ||
          (current.fit === target.fit &&
            (current.minZoom ?? null) === (target.minZoom ?? null)))
      );
    }
    if (row.zoom !== undefined) {
      return (mode?.type ?? 'fixed') === 'fixed' && Math.abs(this.store.zoom() - row.zoom) < 0.001;
    }
    return false;
  }

  /** The next rung, exposed for the shell's keyboard shortcuts. */
  step(direction: 'in' | 'out'): void {
    const next = stepZoomLevel(this.store.zoom(), direction);
    if (next !== null) this.store.setZoom(next);
  }
}
