import { TestBed } from '@angular/core/testing';
import { beforeEach, describe, expect, it } from 'vitest';
import { DocxToolbarComponent } from './docx-toolbar';
import { DocxEditorStore } from '../../core/docx-editor-store';
import { provideDocxEditor } from '../../core/provide-docx-editor';

/**
 * The toolbar is DERIVED from the engine registry, so these assertions are about the
 * projection: every default group's controls appear, each carries its slot id, and enabled
 * state comes from the engine (`editor is not ready` before a document is mounted).
 */
describe('DocxToolbarComponent', () => {
  beforeEach(() => {
    TestBed.resetTestingModule();
    TestBed.configureTestingModule({
      imports: [DocxToolbarComponent],
      providers: [provideDocxEditor({})],
    });
  });

  function render(): HTMLElement {
    const fixture = TestBed.createComponent(DocxToolbarComponent);
    fixture.detectChanges();
    return fixture.nativeElement as HTMLElement;
  }

  it('renders the registry’s default bar', () => {
    const element = render();
    const slots = [...element.querySelectorAll('[data-slot]')].map((node) =>
      node.getAttribute('data-slot')
    );
    expect(slots.length).toBeGreaterThanOrEqual(20);
    expect(new Set(slots).size).toBe(slots.length);
    expect(slots).toContain('text.bold');
    expect(slots).toContain('history.undo');
    expect(slots).toContain('alignment.left');
    expect(slots).toContain('list.bullet');
  });

  it('labels controls from the catalogue', () => {
    const element = render();
    const bold = element.querySelector('[data-slot="text.bold"]');
    expect(bold?.getAttribute('aria-label')).toBe('Bold (Ctrl+B)');
    expect(bold?.getAttribute('aria-pressed')).toBe('false');
  });

  it('disables every command before a document is mounted, quoting the engine’s reason', () => {
    const element = render();
    const bold = element.querySelector('[data-slot="text.bold"]') as HTMLButtonElement;
    expect(bold.disabled).toBe(true);
    expect(bold.getAttribute('title')).toContain('Bold (Ctrl+B)');
    // The engine's own refusal, localized — not an adapter-invented string.
    expect(bold.getAttribute('title')).toContain('No document is loaded');
  });

  it('renders the pickers as their own controls, not as icon buttons', () => {
    const element = render();
    expect(element.querySelector('.docx-toolbar__font-family-trigger')).not.toBeNull();
    expect(element.querySelector('.docx-toolbar__style-trigger')).not.toBeNull();
    expect(element.querySelector('.docx-toolbar__zoom')).not.toBeNull();
    expect(element.querySelector('.docx-toolbar__colorsplit')).not.toBeNull();
    expect(element.querySelector('.docx-toolbar__mode')).not.toBeNull();
    // The style picker shows the registry's placeholder before a document is loaded.
    expect(element.querySelector('.docx-toolbar__style-trigger')?.textContent?.trim()).toBe(
      'Normal text'
    );
  });

  it('omits the slots this adapter has no chrome for', () => {
    const element = render();
    expect(element.querySelector('[data-slot="file.pageSetup"]')).toBeNull();
    expect(element.querySelector('[data-slot="paragraph.dialog"]')).toBeNull();
    expect(element.querySelector('[data-slot="image.properties"]')).toBeNull();
  });

  it('separates registry groups', () => {
    const element = render();
    const groups = element.querySelectorAll('.docx-toolbar__group');
    expect(groups.length).toBeGreaterThanOrEqual(8);
    expect(element.querySelectorAll('.docx-toolbar__separator').length).toBe(groups.length);
  });

  it('is a toolbar for assistive technology', () => {
    const element = render();
    const toolbar = element.querySelector('.docx-toolbar');
    expect(toolbar?.getAttribute('role')).toBe('toolbar');
    expect(toolbar?.getAttribute('aria-label')).toBe('Formatting toolbar');
  });
});

describe('DocxToolbarComponent with a host save handler', () => {
  it('enables the history slots only once the engine is ready', () => {
    TestBed.resetTestingModule();
    TestBed.configureTestingModule({
      imports: [DocxToolbarComponent],
      providers: [provideDocxEditor({})],
    });
    const store = TestBed.inject(DocxEditorStore);
    const fixture = TestBed.createComponent(DocxToolbarComponent);
    fixture.detectChanges();
    const undo = (fixture.nativeElement as HTMLElement).querySelector(
      '[data-slot="history.undo"]'
    ) as HTMLButtonElement;
    expect(undo.disabled).toBe(true);
    expect(store.canUndo()).toBe(false);
  });
});
