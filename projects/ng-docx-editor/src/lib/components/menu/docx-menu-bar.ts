import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core';
import { CHROME_MENUS } from '@docx-editor.dev/core/editor';
import type {
  ChromeMenu,
  ChromeMenuEntry,
  ChromeMenuItemEntry,
  ChromeMenuSubmenuEntry,
  ChromeSlotId,
} from '@docx-editor.dev/core/editor';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { controlTitle, refusalText, slotLabel } from '../../core/labels';
import { slotIconPaths } from '../../core/chrome-registry';
import { isAdapterPending } from '../../core/adapter-pending';
import { DocxIconComponent } from '../shared/docx-icon';
import { DocxMenuDirective } from '../shared/docx-menu.directive';
import { DocxTableGridComponent } from './docx-table-grid';

/**
 * The packaged menu bar — File · Format · Insert · Review · Help — derived FROM the engine's
 * `CHROME_MENUS`.
 *
 * Every actionable row is a chrome slot, so a menu row and its toolbar twin share one label,
 * one icon, one command and one enabled state: the drift that comes from describing "Bold"
 * twice cannot happen here. Submenus are the registry's two-level flyouts; the insert-table
 * row opens the grid picker because the size is a payload the engine cannot invent.
 */
@Component({
  selector: 'docx-menu-bar',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxIconComponent, DocxMenuDirective, DocxTableGridComponent],
  template: `
    <div
      class="docx-menubar"
      role="menubar"
      [attr.aria-label]="t()('titleBar.menuBarAriaLabel')"
    >
      @for (menu of menus; track menu.id) {
        <div class="docx-menubar__menu-root" docxMenu #root="docxMenu">
          <button
            type="button"
            class="docx-menubar__trigger"
            data-menu-trigger
            role="menuitem"
            [attr.aria-haspopup]="true"
            [attr.aria-expanded]="root.open()"
            (click)="root.toggle()"
          >
            {{ t()(menu.labelKey) }}
          </button>
          @if (root.open()) {
            <div class="docx-menubar__menu" role="menu">
              @for (entry of menu.entries; track $index) {
                @switch (entry.kind) {
                  @case ('separator') {
                    <div class="docx-menubar__group" role="separator"></div>
                  }
                  @case ('item') {
                    @if (!isPending(entry.slot)) {
                      <button
                        type="button"
                        class="docx-menubar__item"
                        data-menu-item
                        role="menuitem"
                        [attr.data-slot]="entry.slot"
                        [attr.aria-disabled]="stateOf(entry.slot).enabled ? null : 'true'"
                        [attr.title]="titleOf(entry.slot)"
                        (click)="runItem(entry)"
                      >
                        <span class="docx-menubar__item-icon">
                          <docx-icon [paths]="pathsOf(entry.slot)" [size]="16" />
                        </span>
                        <span class="docx-menubar__item-label">{{ labelOf(entry) }}</span>
                        @if (entry.shortcutKey) {
                          <span class="docx-menubar__item-shortcut">{{ t()(entry.shortcutKey) }}</span>
                        }
                      </button>
                      @if (entry.picker === 'tableGrid' && gridOpen()) {
                        <docx-table-grid (pick)="insertTable($event)" />
                      }
                    }
                  }
                  @case ('submenu') {
                    @if (!submenuPending(entry)) {
                      <div class="docx-menubar__submenu" role="group">
                        <div class="docx-menubar__item docx-menubar__item--parent">
                          <span class="docx-menubar__item-icon">
                            <docx-icon [paths]="entry.paths" [size]="16" />
                          </span>
                          <span class="docx-menubar__item-label">{{ t()(entry.labelKey) }}</span>
                          <span class="docx-menubar__item-caret">▸</span>
                        </div>
                        @for (child of entry.items; track $index) {
                          @if (child.kind === 'item' && !isPending(child.slot)) {
                            <button
                              type="button"
                              class="docx-menubar__item docx-menubar__item--nested"
                              data-menu-item
                              role="menuitem"
                              [attr.data-slot]="child.slot"
                              [attr.aria-disabled]="stateOf(child.slot).enabled ? null : 'true'"
                              [attr.title]="titleOf(child.slot)"
                              (click)="runItem(child)"
                            >
                              <span class="docx-menubar__item-icon">
                                <docx-icon [paths]="pathsOf(child.slot)" [size]="16" />
                              </span>
                              <span class="docx-menubar__item-label">{{ labelOf(child) }}</span>
                            </button>
                          }
                        }
                      </div>
                    }
                  }
                }
              }
            </div>
          }
        </div>
      }
    </div>
  `,
  styles: [
    `
      :host {
        display: block;
      }
      .docx-menubar {
        display: flex;
        align-items: center;
        gap: 2px;
      }
      .docx-menubar__menu-root {
        position: relative;
      }
      .docx-menubar__trigger {
        border: none;
        background: transparent;
        color: var(--doc-text);
        font: inherit;
        font-size: 13px;
        padding: 4px 8px;
        border-radius: 4px;
        cursor: pointer;
      }
      .docx-menubar__trigger:hover,
      .docx-menubar__menu-root[data-menu-open] .docx-menubar__trigger {
        background: var(--doc-bg-hover);
      }
      .docx-menubar__menu {
        position: absolute;
        top: 100%;
        inset-inline-start: 0;
        z-index: var(--doc-z-overlay);
        margin-top: 4px;
        min-width: 240px;
        padding: 4px;
        background: var(--doc-surface);
        border: 1px solid var(--doc-border);
        border-radius: 8px;
        box-shadow: var(--doc-shadow-lg);
      }
      .docx-menubar__item {
        display: flex;
        align-items: center;
        gap: 8px;
        width: 100%;
        padding: 7px 10px;
        border: none;
        border-radius: 4px;
        background: transparent;
        color: var(--doc-text);
        font: inherit;
        font-size: 13px;
        text-align: start;
        cursor: pointer;
        white-space: nowrap;
      }
      .docx-menubar__item:hover {
        background: var(--doc-bg-hover);
      }
      .docx-menubar__item[aria-disabled='true'] {
        opacity: 0.5;
        cursor: default;
      }
      .docx-menubar__item-icon {
        display: inline-flex;
        width: 16px;
        flex: none;
      }
      .docx-menubar__item-label {
        flex: 1;
      }
      .docx-menubar__item-shortcut {
        color: var(--doc-text-muted);
        font-size: 11px;
      }
      .docx-menubar__group {
        height: 1px;
        margin: 4px 0;
        background: var(--doc-border);
      }
      .docx-menubar__submenu {
        padding: 2px 0;
      }
      .docx-menubar__item--parent {
        color: var(--doc-text-muted);
        cursor: default;
      }
      .docx-menubar__item--nested {
        padding-inline-start: 26px;
      }
      .docx-menubar__grid {
        padding: 6px 10px 10px;
      }
      .docx-menubar__grid-caption {
        display: block;
        margin-bottom: 4px;
        color: var(--doc-text-muted);
        font-size: 11px;
      }
      .docx-menubar__grid-row {
        display: flex;
      }
      .docx-menubar__grid-cell {
        width: 18px;
        height: 18px;
        margin: 1px;
        padding: 0;
        border: 1px solid var(--doc-border);
        border-radius: 2px;
        background: var(--doc-surface);
        cursor: pointer;
      }
      .docx-menubar__grid-cell[data-active] {
        background: var(--doc-accent);
        border-color: var(--doc-accent);
      }
    `,
  ],
})
export class DocxMenuBarComponent {
  private readonly store = inject(DocxEditorStore);

  readonly t = this.store.t;
  /** The menu registry, in bar order. */
  readonly menus: readonly ChromeMenu[] = CHROME_MENUS;
  /** Whether the insert-table grid is showing under its row. */
  readonly gridOpen = signal(false);

  readonly states = computed(() => this.store.menuCommandStates());

  stateOf(slot: ChromeSlotId) {
    return this.states().get(slot) ?? { id: slot, enabled: false, disabledReason: null, active: false };
  }

  labelOf(entry: ChromeMenuItemEntry): string {
    return entry.labelKey ? this.t()(entry.labelKey) : slotLabel(this.t(), entry.slot);
  }

  titleOf(slot: ChromeSlotId): string {
    return controlTitle(this.t(), slot, this.stateOf(slot));
  }

  pathsOf(slot: ChromeSlotId): readonly string[] | null {
    return slotIconPaths(slot);
  }

  isPending(slot: ChromeSlotId): boolean {
    return isAdapterPending(slot);
  }

  submenuPending(entry: ChromeMenuSubmenuEntry): boolean {
    return entry.items.every((item) => item.kind === 'item' && isAdapterPending(item.slot));
  }

  runItem(entry: ChromeMenuItemEntry): void {
    const state = this.stateOf(entry.slot);
    if (!state.enabled) {
      const reason = refusalText(this.t(), state.disabledReason);
      if (reason) this.store.notice.set(reason);
      return;
    }
    if (entry.picker === 'tableGrid') {
      this.gridOpen.update((value) => !value);
      return;
    }
    this.store.run(entry.slot);
  }

  insertTable(size: { rows: number; cols: number }): void {
    this.gridOpen.set(false);
    this.store.insertTable(size.rows, size.cols);
  }
}

export type { ChromeMenuEntry };
