import { Directive, ElementRef, effect, inject, signal } from '@angular/core';

/**
 * The behaviour behind every packaged dropdown: an open flag, outside-press and Escape
 * dismissal, and "picking a row closes it".
 *
 * Used as `docxMenu` on a positioning wrapper that contains both the trigger and the panel,
 * so containment answers "did the press land inside this control". Rows in the panel carry
 * `data-menu-item`; a click on one closes the panel, while a click on the trigger
 * (`data-menu-trigger`) is left to the trigger's own `toggle()`.
 *
 * @example
 * ```html
 * <div class="docx-toolbar__font-family" docxMenu #menu="docxMenu">
 *   <button data-menu-trigger [attr.aria-expanded]="menu.open()" (click)="menu.toggle()">Arial</button>
 *   @if (menu.open()) {
 *     <div class="docx-toolbar__menu">
 *       <button data-menu-item (click)="pick('Arial')">Arial</button>
 *     </div>
 *   }
 * </div>
 * ```
 */
@Directive({
  selector: '[docxMenu]',
  exportAs: 'docxMenu',
  host: {
    '[attr.data-menu-open]': 'open() ? "" : null',
  },
})
export class DocxMenuDirective {
  private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);

  /** Whether the panel is showing. */
  readonly open = signal(false);

  constructor() {
    effect((onCleanup) => {
      if (!this.open()) return;
      const doc = this.host.nativeElement.ownerDocument;
      const onPointerDown = (event: Event) => {
        const target = event.target;
        if (target instanceof Node && !this.host.nativeElement.contains(target)) this.close();
      };
      const onClick = (event: Event) => {
        const target = event.target;
        if (!(target instanceof Element)) return;
        if (target.closest('[data-menu-item]')) this.close();
      };
      const onKeyDown = (event: KeyboardEvent) => {
        if (event.key === 'Escape') {
          event.stopPropagation();
          this.close();
          this.focusTrigger();
        }
      };
      doc.addEventListener('pointerdown', onPointerDown, true);
      doc.addEventListener('click', onClick, true);
      doc.addEventListener('keydown', onKeyDown, true);
      onCleanup(() => {
        doc.removeEventListener('pointerdown', onPointerDown, true);
        doc.removeEventListener('click', onClick, true);
        doc.removeEventListener('keydown', onKeyDown, true);
      });
    });
  }

  /** Show the panel. */
  show(): void {
    this.open.set(true);
  }

  /** Hide the panel. */
  close(): void {
    this.open.set(false);
  }

  /** Toggle the panel. */
  toggle(): void {
    this.open.update((value) => !value);
  }

  /** Return focus to the trigger, so keyboard users keep their place after Escape. */
  private focusTrigger(): void {
    const trigger = this.host.nativeElement.querySelector<HTMLElement>('[data-menu-trigger]');
    trigger?.focus();
  }
}
