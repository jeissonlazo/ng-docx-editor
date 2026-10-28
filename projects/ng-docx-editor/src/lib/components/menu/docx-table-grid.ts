import { ChangeDetectionStrategy, Component, computed, output, signal } from '@angular/core';

/** The largest grid the insert-table picker offers, matching Word's 8 × 8 affordance. */
const GRID = 8;

/**
 * Word's insert-table grid: hover `rows × cols`, click to insert.
 *
 * It exists because the slot that opens it (`table.insert`) needs a size the engine cannot
 * invent — the row is enabled by the engine's own probe and this grid supplies the payload.
 */
@Component({
  selector: 'docx-table-grid',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="docx-menubar__grid" role="group">
      <span class="docx-menubar__grid-caption">{{ caption() }}</span>
      <div class="docx-menubar__grid-cells">
        @for (row of rows; track row) {
          <div class="docx-menubar__grid-row">
            @for (col of rows; track col) {
              <button
                type="button"
                class="docx-menubar__grid-cell"
                [attr.data-active]="col <= hoverCols() && row <= hoverRows() ? '' : null"
                [attr.aria-label]="row + ' × ' + col"
                (mouseenter)="hover(row, col)"
                (focus)="hover(row, col)"
                (click)="pick.emit({ rows: row, cols: col })"
              ></button>
            }
          </div>
        }
      </div>
    </div>
  `,
})
export class DocxTableGridComponent {
  /** Emits the picked size. */
  readonly pick = output<{ rows: number; cols: number }>();

  readonly rows = Array.from({ length: GRID }, (_, index) => index + 1);
  readonly hoverRows = signal(1);
  readonly hoverCols = signal(1);
  readonly caption = computed(() => `${this.hoverRows()} × ${this.hoverCols()}`);

  hover(rows: number, cols: number): void {
    this.hoverRows.set(rows);
    this.hoverCols.set(cols);
  }
}
