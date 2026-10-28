import { ChangeDetectionStrategy, Component, computed, inject, input, output } from '@angular/core';
import { chromeSlotIsToggle } from '@docx-editor.dev/core/editor';
import type { ChromeSlotId } from '@docx-editor.dev/core/editor';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { controlTitle, slotLabel } from '../../core/labels';
import { slotIconPaths } from '../../core/chrome-registry';
import { DocxIconComponent } from '../shared/docx-icon';

/**
 * One registry-driven toolbar button.
 *
 * Enabled state, pressed state and the refusal message all come from the engine through
 * `toolbarCommandState`, so a button and the command behind it cannot disagree. The icon is
 * the registry's own path data.
 */
@Component({
  selector: 'docx-toolbar-button',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxIconComponent],
  template: `
    <button
      type="button"
      class="docx-toolbar__button"
      [attr.data-slot]="slot()"
      [attr.data-active]="state().active ? '' : null"
      [attr.data-value]="state().value ?? null"
      [attr.aria-pressed]="isToggle() ? state().active : null"
      [attr.aria-label]="label()"
      [attr.title]="title()"
      [disabled]="!state().enabled"
      (mousedown)="$event.preventDefault()"
      (click)="run()"
    >
      <docx-icon [paths]="paths()" />
    </button>
  `,
})
export class DocxToolbarButtonComponent {
  private readonly store = inject(DocxEditorStore);

  /** The chrome slot this button drives. */
  readonly slot = input.required<ChromeSlotId>();
  /** Overrides the registry glyph, for a host-composed button. */
  readonly icon = input<readonly string[] | null | undefined>(undefined);
  /**
   * Hand the CLICK to the host instead of running the slot's command — for a slot whose
   * payload comes from chrome the engine cannot supply (an image from a file picker).
   */
  readonly delegate = input(false);
  /** Fired when `delegate` is set and the button is pressed. */
  readonly activate = output<void>();

  readonly state = computed(() => this.store.commandState(this.slot()));
  readonly label = computed(() => slotLabel(this.store.t(), this.slot()));
  readonly title = computed(() => controlTitle(this.store.t(), this.slot(), this.state()));
  readonly paths = computed(() => this.icon() ?? slotIconPaths(this.slot()));
  readonly isToggle = computed(() => chromeSlotIsToggle(this.slot()));

  run(): void {
    if (this.delegate()) this.activate.emit();
    else this.store.run(this.slot());
  }
}
