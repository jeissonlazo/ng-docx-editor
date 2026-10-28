import { ChangeDetectionStrategy, Component, computed, inject, input, signal } from '@angular/core';
import type { ChromeSlotId } from '@docx-editor.dev/core/editor';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { refusalText, slotLabel } from '../../core/labels';
import { slotIconPaths } from '../../core/chrome-registry';
import { DocxIconComponent } from '../shared/docx-icon';
import { DocxMenuDirective } from '../shared/docx-menu.directive';

/** Word's standard-colour row, as `w:color` hex values. */
const STANDARD_COLORS: readonly string[] = [
  'C00000',
  'FF0000',
  'FFC000',
  'FFFF00',
  '92D050',
  '00B050',
  '00B0F0',
  '0070C0',
  '002060',
  '7030A0',
];

/** `ST_HighlightColor` names with the hexes they paint. */
const HIGHLIGHT_COLORS: readonly { readonly value: string; readonly hex: string }[] = [
  { value: 'yellow', hex: 'FFFF00' },
  { value: 'green', hex: '00FF00' },
  { value: 'cyan', hex: '00FFFF' },
  { value: 'magenta', hex: 'FF00FF' },
  { value: 'blue', hex: '0000FF' },
  { value: 'red', hex: 'FF0000' },
  { value: 'darkBlue', hex: '00008B' },
  { value: 'darkCyan', hex: '008080' },
  { value: 'darkGreen', hex: '008000' },
  { value: 'darkMagenta', hex: '800080' },
  { value: 'darkRed', hex: '8B0000' },
  { value: 'darkYellow', hex: '808000' },
  { value: 'darkGray', hex: '808080' },
  { value: 'lightGray', hex: 'C0C0C0' },
  { value: 'black', hex: '000000' },
  { value: 'white', hex: 'FFFFFF' },
];

/** Office's default scheme, for a document that declares none. */
const DEFAULT_THEME_HEXES: readonly string[] = [
  'FFFFFF',
  '000000',
  'E7E6E6',
  '44546A',
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
];

const THEME_SLOTS = [
  'background1',
  'text1',
  'background2',
  'text2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
] as const;

/**
 * The split colour control: the main half applies the last-used colour, the caret opens
 * Word's picker (clear row, theme matrix, standard row, custom hex).
 *
 * One component serves both colour slots: `text.color` dispatches six-digit hex values and
 * `text.highlight` dispatches the `ST_HighlightColor` names the format actually stores.
 */
@Component({
  selector: 'docx-color-split',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxIconComponent, DocxMenuDirective],
  template: `
    <div class="docx-toolbar__colorsplit" docxMenu #menu="docxMenu">
      <button
        type="button"
        class="docx-toolbar__button docx-toolbar__colorsplit-main"
        [attr.data-slot]="slot()"
        [attr.title]="title()"
        [attr.aria-label]="label()"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="apply(lastUsed())"
      >
        <docx-icon [paths]="paths()" />
        <span class="docx-toolbar__colorsplit-bar" [style.background]="barColor()"></span>
      </button>
      <button
        type="button"
        class="docx-toolbar__colorsplit-caret"
        data-menu-trigger
        [attr.aria-label]="label()"
        [attr.aria-expanded]="menu.open()"
        [attr.title]="title()"
        [disabled]="!state().enabled"
        (mousedown)="$event.preventDefault()"
        (click)="menu.toggle()"
      >
        ▾
      </button>

      @if (menu.open()) {
        <div class="docx-toolbar__swatch-popup">
          <div class="docx-toolbar__swatch-section">
            <button type="button" class="docx-toolbar__swatch-clear" data-menu-item (click)="clear()">
              <span
                class="docx-toolbar__swatch-clear-chip"
                [class.docx-toolbar__swatch-clear-chip--none]="isHighlight() || slot() === 'text.color'"
              ></span>
              {{ isHighlight() ? t()('colorPicker.noColor') : t()('colorPicker.automatic') }}
            </button>
          </div>

          <div class="docx-toolbar__swatch-section">
            <span class="docx-toolbar__swatch-heading">{{ t()('colorPicker.themeColors') }}</span>
            <div class="docx-toolbar__swatch-grid docx-toolbar__swatch-grid--theme">
              @for (color of themeColors(); track color.slot) {
                <button
                  type="button"
                  class="docx-toolbar__swatch"
                  data-menu-item
                  [attr.data-light]="isLight(color.hex) ? '' : null"
                  [attr.title]="t()('colorPicker.theme.' + color.slot)"
                  [attr.aria-label]="t()('colorPicker.theme.' + color.slot)"
                  [style.background]="'#' + color.hex"
                  (click)="apply(color.value)"
                ></button>
              }
            </div>
          </div>

          @if (isHighlight()) {
            <div class="docx-toolbar__swatch-section">
              <span class="docx-toolbar__swatch-heading">{{ t()('colorPicker.highlightColors') }}</span>
              <div class="docx-toolbar__swatch-grid docx-toolbar__swatch-grid--highlight">
                @for (color of highlightColors; track color.value) {
                  <button
                    type="button"
                    class="docx-toolbar__swatch"
                    data-menu-item
                    [attr.data-light]="isLight(color.hex) ? '' : null"
                    [attr.title]="color.value"
                    [attr.aria-label]="color.value"
                    [style.background]="'#' + color.hex"
                    (click)="apply(color.value)"
                  ></button>
                }
              </div>
            </div>
          } @else {
            <div class="docx-toolbar__swatch-section">
              <span class="docx-toolbar__swatch-heading">{{ t()('colorPicker.standardColors') }}</span>
              <div class="docx-toolbar__swatch-grid">
                @for (hex of standardColors; track hex) {
                  <button
                    type="button"
                    class="docx-toolbar__swatch"
                    data-menu-item
                    [attr.data-light]="isLight(hex) ? '' : null"
                    [attr.title]="'#' + hex"
                    [attr.aria-label]="'#' + hex"
                    [style.background]="'#' + hex"
                    (click)="apply(hex)"
                  ></button>
                }
              </div>
            </div>

            <div class="docx-toolbar__swatch-section">
              <span class="docx-toolbar__swatch-heading">{{ t()('colorPicker.customColor') }}</span>
              <div class="docx-toolbar__swatch-custom">
                <span class="docx-toolbar__swatch-hash">#</span>
                <input
                  class="docx-toolbar__swatch-hex"
                  type="text"
                  maxlength="6"
                  spellcheck="false"
                  [value]="hex()"
                  (input)="hex.set($any($event.target).value)"
                  (keydown.enter)="applyHex()"
                />
                <button type="button" class="docx-toolbar__swatch-apply" (click)="applyHex()">
                  {{ t()('colorPicker.apply') }}
                </button>
              </div>
            </div>
          }
        </div>
      }
    </div>
  `,
})
export class DocxColorSplitComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  /** Which colour slot this control drives. */
  readonly slot = input.required<'text.color' | 'text.highlight'>();

  readonly standardColors = STANDARD_COLORS;
  readonly highlightColors = HIGHLIGHT_COLORS;
  readonly hex = signal('');

  private readonly lastPicked = signal<string | null>(null);

  readonly isHighlight = computed(() => this.slot() === 'text.highlight');
  readonly state = computed(() => this.store.commandState(this.slot() as ChromeSlotId));
  readonly label = computed(() => slotLabel(this.t(), this.slot() as ChromeSlotId));
  readonly title = computed(() => {
    const reason = this.state().enabled ? null : refusalText(this.t(), this.state().disabledReason);
    return reason ? `${this.label()} — ${reason}` : this.label();
  });
  readonly paths = computed(() => slotIconPaths(this.slot() as ChromeSlotId));

  /** The colour in force at the selection, as the slot stores it. */
  readonly currentValue = computed(() => {
    const formatting = this.store.formatting();
    if (!formatting) return null;
    return this.isHighlight()
      ? (formatting.highlight ?? null)
      : formatting.color?.kind === 'hex'
        ? formatting.color.value.replace(/^#/, '')
        : null;
  });

  /** What the main half applies: the last pick, else the current value, else Word's default. */
  readonly lastUsed = computed(
    () => this.lastPicked() ?? this.currentValue() ?? (this.isHighlight() ? 'yellow' : 'C00000')
  );

  readonly barColor = computed(() => {
    const value = this.lastUsed();
    if (!value) return 'transparent';
    const match = HIGHLIGHT_COLORS.find((color) => color.value === value);
    return `#${match?.hex ?? value.replace(/^#/, '')}`;
  });

  readonly themeColors = computed(() => {
    const declared = this.store.themeColors();
    if (declared.length >= 10) {
      return declared.slice(0, 10).map((entry, index) => ({
        slot: THEME_SLOTS[index] ?? entry.slot,
        hex: entry.hex,
        value: entry.hex,
      }));
    }
    return DEFAULT_THEME_HEXES.map((hex, index) => ({
      slot: THEME_SLOTS[index]!,
      hex,
      value: hex,
    }));
  });

  apply(value: string | null): void {
    if (value === null) {
      this.store.run(this.slot(), this.isHighlight() ? 'none' : 'auto');
      return;
    }
    this.lastPicked.set(value);
    this.store.run(this.slot(), value);
  }

  clear(): void {
    this.lastPicked.set(null);
    this.store.run(this.slot(), this.isHighlight() ? 'none' : 'auto');
  }

  applyHex(): void {
    const hex = this.hex().replace(/^#/, '').trim();
    if (!/^[0-9a-fA-F]{6}$/.test(hex)) return;
    this.apply(hex.toUpperCase());
  }

  /** Whether a swatch needs its hairline border to stay visible on a white surface. */
  isLight(hex: string): boolean {
    const value = hex.replace(/^#/, '');
    if (value.length !== 6) return false;
    const r = parseInt(value.slice(0, 2), 16);
    const g = parseInt(value.slice(2, 4), 16);
    const b = parseInt(value.slice(4, 6), 16);
    return 0.299 * r + 0.587 * g + 0.114 * b > 200;
  }
}
