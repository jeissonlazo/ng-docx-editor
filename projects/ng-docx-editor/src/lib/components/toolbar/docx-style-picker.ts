import { ChangeDetectionStrategy, Component, computed, inject } from '@angular/core';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { refusalText, slotLabel } from '../../core/labels';
import { DocxMenuDirective } from '../shared/docx-menu.directive';

/**
 * The paragraph-style picker: every style the document defines, each row drawn in its own
 * face so the list reads as a preview rather than a list of identical labels.
 *
 * The value is the style id; the engine refuses an id the document does not define, so the
 * list comes from the document itself (`getDocumentStyles`).
 */
@Component({
  selector: 'docx-style-picker',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuDirective],
  template: `
    <div class="docx-toolbar__style" docxMenu #menu="docxMenu">
      <button
        type="button"
        class="docx-toolbar__style-trigger"
        data-menu-trigger
        [attr.aria-label]="t()('styles.selectAriaLabel')"
        [attr.aria-expanded]="menu.open()"
        [attr.title]="title()"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="menu.toggle()"
      >
        <span>{{ currentName() }}</span>
      </button>
      @if (menu.open()) {
        <div class="docx-toolbar__menu docx-toolbar__style-content">
          @for (style of paragraphStyles(); track style.styleId) {
            <button
              type="button"
              class="docx-toolbar__style-item"
              data-menu-item
              [style.font-family]="style.preview.fontFamily ?? null"
              [style.font-size.px]="previewSize(style.preview.fontSizePt)"
              [style.font-weight]="style.preview.bold ? 700 : 400"
              [style.font-style]="style.preview.italic ? 'italic' : null"
              (click)="pick(style.styleId)"
            >
              {{ style.name }}
              @if (style.styleId === currentId()) {
                <span class="docx-toolbar__menu-check">✓</span>
              }
            </button>
          } @empty {
            <div class="docx-toolbar__menu-label">{{ t()('styles.normalText') }}</div>
          }
        </div>
      }
    </div>
  `,
})
export class DocxStylePickerComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  readonly slot = 'styles.style' as const;
  readonly state = computed(() => this.store.commandState(this.slot));
  readonly label = computed(() => slotLabel(this.store.t(), this.slot));
  readonly title = computed(() => {
    const reason = this.state().enabled ? null : refusalText(this.t(), this.state().disabledReason);
    return reason ? `${this.label()} — ${reason}` : this.label();
  });

  readonly currentId = computed(() => this.store.formatting()?.styleId ?? null);
  readonly paragraphStyles = computed(() =>
    this.store.documentStyles().filter((style) => style.type === 'paragraph')
  );
  readonly currentName = computed(() => {
    const id = this.currentId();
    const match = this.paragraphStyles().find((style) => style.styleId === id);
    return match?.name ?? this.t()('styles.normalText');
  });

  pick(styleId: string): void {
    this.store.setParagraphStyle(styleId);
  }

  /** Preview size: the style's own point size, capped so the list stays one line per row. */
  previewSize(points: number | null): number | null {
    if (points === null) return null;
    return Math.min(Math.max(points, 9), 18);
  }
}
