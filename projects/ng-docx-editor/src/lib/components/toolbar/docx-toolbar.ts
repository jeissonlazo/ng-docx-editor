import { ChangeDetectionStrategy, Component, computed, inject, input } from '@angular/core';
import { formattingBarChromeGroups } from '@docx-editor.dev/core/editor';
import type { ChromeControl, ChromeGroup, ChromeSlotId } from '@docx-editor.dev/core/editor';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { isAdapterPending } from '../../core/adapter-pending';
import { DocxToolbarButtonComponent } from './docx-toolbar-button';
import { DocxFontFamilyComponent } from './docx-font-family';
import { DocxFontSizeComponent } from './docx-font-size';
import { DocxStylePickerComponent } from './docx-style-picker';
import { DocxColorSplitComponent } from './docx-color-split';
import { DocxLineSpacingComponent } from './docx-line-spacing';
import { DocxZoomStepperComponent } from './docx-zoom-stepper';
import { DocxModePillComponent } from './docx-mode-pill';

/**
 * The default formatting bar, derived from the engine's chrome registry.
 *
 * Controls are never hand-listed: the bar walks `formattingBarChromeGroups()` (the registry's
 * default groups, plus the contextual image group when a drawing is selected, plus text
 * direction) and renders each slot with the shape the registry declares. A slot the registry
 * adds appears here without this file being edited, and every control's enabled and pressed
 * state comes from the engine.
 */
@Component({
  selector: 'docx-toolbar',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [
    DocxToolbarButtonComponent,
    DocxFontFamilyComponent,
    DocxFontSizeComponent,
    DocxStylePickerComponent,
    DocxColorSplitComponent,
    DocxLineSpacingComponent,
    DocxZoomStepperComponent,
    DocxModePillComponent,
  ],
  template: `
    <div class="docx-toolbar" role="toolbar" [attr.aria-label]="t()('toolbar.ariaLabel')">
      @for (group of groups(); track group.id) {
        <div class="docx-toolbar__group" [attr.data-group]="group.id">
          @for (control of group.controls; track control.id) {
            @if (!isPending(slotOf(group, control))) {
              @switch (slotOf(group, control)) {
                @case ('zoom.level') {
                  <docx-zoom-stepper />
                }
                @case ('styles.style') {
                  <docx-style-picker />
                }
                @case ('font.family') {
                  <docx-font-family />
                }
                @case ('font.size') {
                  <docx-font-size />
                }
                @case ('text.color') {
                  <docx-color-split slot="text.color" />
                }
                @case ('text.highlight') {
                  <docx-color-split slot="text.highlight" />
                }
                @case ('list.lineSpacing') {
                  <docx-line-spacing />
                }
                @case ('review.editingMode') {
                  <docx-mode-pill />
                }
                @case ('image.insert') {
                  <input
                    #imageFile
                    type="file"
                    accept="image/*"
                    hidden
                    (change)="insertImage($event)"
                  />
                  <docx-toolbar-button
                    slot="image.insert"
                    [delegate]="true"
                    (activate)="imageFile.click()"
                  />
                }
                @default {
                  <docx-toolbar-button [slot]="slotOf(group, control)" />
                }
              }
            }
          }
          <span class="docx-toolbar__separator" aria-hidden="true"></span>
        </div>
      }
    </div>
  `,
})
export class DocxToolbarComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;

  /** Extra controls the host wants in the bar, by slot id. */
  readonly extraSlots = input<readonly ChromeSlotId[]>([]);

  readonly groups = computed<readonly ChromeGroup[]>(() =>
    formattingBarChromeGroups(this.store.snapshot().image, { direction: true })
  );

  slotOf(group: ChromeGroup, control: ChromeControl): ChromeSlotId {
    return `${group.id}.${control.id}` as ChromeSlotId;
  }

  isPending(slot: ChromeSlotId): boolean {
    return isAdapterPending(slot);
  }

  async insertImage(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) await this.store.insertImageFile(file);
  }
}
