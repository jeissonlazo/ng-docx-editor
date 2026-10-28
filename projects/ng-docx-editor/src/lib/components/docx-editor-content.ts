import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  ElementRef,
  afterNextRender,
  inject,
  viewChild,
} from '@angular/core';
import { DocxEditorStore } from '../core/docx-editor-store';

/**
 * The element the engine paints pages into — the attach point.
 *
 * The engine owns everything inside this div. `attach` runs after the first render (the
 * element must already be connected, because the engine discovers its scroller by walking
 * ancestors), and `detach` on destroy parks the editor instead of destroying it, so moving
 * the mount point elsewhere keeps the undo history, selection and reading position.
 *
 * Paste and drop of image files route through the shared image-insert path; the engine's own
 * paste handling stands down when the payload carries content it will land itself (Word on
 * macOS pastes a rendered PNG beside copied text).
 */
@Component({
  selector: 'docx-editor-content',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div
      #surface
      class="docx-paginated-surface"
      (paste)="onPaste($event)"
      (dragover)="onDragOver($event)"
      (drop)="onDrop($event)"
    ></div>
  `,
  styles: [
    `
      :host {
        display: contents;
      }
    `,
  ],
})
export class DocxEditorContentComponent {
  private readonly store = inject(DocxEditorStore);
  private readonly destroyRef = inject(DestroyRef);
  private readonly surface = viewChild.required<ElementRef<HTMLDivElement>>('surface');

  constructor() {
    afterNextRender(() => {
      this.store.attach(this.surface().nativeElement);
    });
    this.destroyRef.onDestroy(() => this.store.detach());
  }

  onPaste(event: ClipboardEvent): void {
    const items = event.clipboardData;
    if (!items) return;
    const file = imageFileOf(items);
    if (!file) return;
    event.preventDefault();
    void this.store.insertImageFile(file);
  }

  onDragOver(event: DragEvent): void {
    if (imageFileOf(event.dataTransfer)) event.preventDefault();
  }

  onDrop(event: DragEvent): void {
    const file = imageFileOf(event.dataTransfer);
    if (!file) return;
    event.preventDefault();
    void this.store.insertImageFile(file);
  }
}

function imageFileOf(transfer: DataTransfer | null): File | null {
  if (!transfer) return null;
  for (const item of Array.from(transfer.items)) {
    if (item.kind !== 'file' || !item.type.startsWith('image/')) continue;
    const file = item.getAsFile();
    if (file) return file;
  }
  return null;
}
