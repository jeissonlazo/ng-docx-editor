import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { refusalText, slotLabel } from '../../core/labels';
import { DocxMenuDirective } from '../shared/docx-menu.directive';

/**
 * The font-family picker: a transparent trigger showing the selection's family, and a
 * searchable list of every family the engine offers.
 *
 * The list is the engine's (`getAvailableFonts` — document fonts plus the configured
 * catalogue), so a family the document names is always pickable even when it is not a
 * standard one.
 */
@Component({
  selector: 'docx-font-family',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuDirective],
  template: `
    <div class="docx-toolbar__font-family" docxMenu #menu="docxMenu">
      <button
        type="button"
        class="docx-toolbar__font-family-trigger"
        data-menu-trigger
        [attr.aria-label]="label()"
        [attr.aria-expanded]="menu.open()"
        [attr.title]="title()"
        [attr.data-disabled]="state().enabled ? null : ''"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="menu.toggle()"
      >
        <span>{{ value() }}</span>
      </button>
      @if (menu.open()) {
        <div class="docx-toolbar__menu docx-toolbar__font-family-content">
          <input
            class="docx-toolbar__font-search"
            type="search"
            autofocus
            [attr.placeholder]="t()('font.search')"
            [attr.aria-label]="t()('font.search')"
            [value]="query()"
            (input)="query.set($any($event.target).value)"
            (keydown.enter)="pickFirst()"
          />
          <div class="docx-toolbar__font-options">
            @for (family of families(); track family) {
              <button
                type="button"
                class="docx-toolbar__font-family-item"
                data-menu-item
                [style.font-family]="family"
                [attr.title]="family"
                (click)="pick(family)"
              >
                {{ family }}
                @if (family === current()) {
                  <span class="docx-toolbar__menu-check">✓</span>
                }
              </button>
            } @empty {
              <div class="docx-toolbar__menu-label">{{ t()('font.documentFonts') }}</div>
            }
          </div>
        </div>
      }
    </div>
  `,
})
export class DocxFontFamilyComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  readonly slot = 'font.family' as const;
  readonly state = computed(() => this.store.commandState(this.slot));
  readonly label = computed(() => slotLabel(this.store.t(), this.slot));
  readonly title = computed(() => {
    const reason = this.state().enabled ? null : refusalText(this.store.t(), this.state().disabledReason);
    return reason ? `${this.label()} — ${reason}` : this.label();
  });
  readonly query = signal('');

  readonly current = computed(() => this.store.formatting()?.fontFamily ?? null);
  readonly value = computed(() => this.current() ?? this.store.t()('font.sansSerif'));

  readonly families = computed(() => {
    const query = this.query().trim().toLowerCase();
    const families = this.store.availableFonts();
    return query ? families.filter((family) => family.toLowerCase().includes(query)) : families;
  });

  pick(family: string): void {
    this.store.setFontFamily(family);
  }

  pickFirst(): void {
    const first = this.families()[0];
    if (first) this.store.setFontFamily(first);
  }
}
