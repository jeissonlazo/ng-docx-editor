import { ChangeDetectionStrategy, Component, computed, input } from '@angular/core';

/**
 * One Material Symbols glyph, rendered as inline SVG from the engine registry's path data.
 *
 * The registry draws in a `0 -960 960 960` viewBox, so the glyph is scaled and offset the way
 * the engine's own chrome does — no icon font, no extra request, no second icon table.
 */
@Component({
  selector: 'docx-icon',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <svg
      [attr.width]="size()"
      [attr.height]="size()"
      viewBox="0 -960 960 960"
      fill="currentColor"
      aria-hidden="true"
      focusable="false"
    >
      @for (path of paths(); track $index) {
        <path [attr.d]="path" />
      }
    </svg>
  `,
  styles: [
    `
      :host {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        line-height: 0;
        flex: none;
      }
    `,
  ],
})
export class DocxIconComponent {
  /** The path data to draw; several paths make one composite glyph. */
  readonly paths = input<readonly string[] | null>(null);
  /** Rendered box in pixels. */
  readonly size = input(18);

  readonly hasPaths = computed(() => (this.paths()?.length ?? 0) > 0);
}
