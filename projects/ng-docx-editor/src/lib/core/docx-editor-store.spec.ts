import { TestBed } from '@angular/core/testing';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { DocxEditorStore } from './docx-editor-store';
import { provideDocxEditor } from './provide-docx-editor';

describe('DocxEditorStore', () => {
  beforeEach(() => {
    TestBed.resetTestingModule();
  });

  it('creates an instance and applies the provided options', () => {
    TestBed.configureTestingModule({ providers: [provideDocxEditor({ author: 'Ada', title: 'Doc' })] });
    const store = TestBed.inject(DocxEditorStore);
    expect(store.options().author).toBe('Ada');
    expect(store.options().title).toBe('Doc');
    expect(store.instance()).not.toBeNull();
    expect(store.getEditor()).not.toBeNull();
  });

  it('is idempotent: an equal patch does not rewrite the options signal', () => {
    TestBed.configureTestingModule({ providers: [provideDocxEditor({ author: 'Ada' })] });
    const store = TestBed.inject(DocxEditorStore);
    const before = store.options();
    store.configure({ author: 'Ada' });
    expect(store.options()).toBe(before);

    store.configure({ author: 'Grace' });
    expect(store.options()).not.toBe(before);
    expect(store.options().author).toBe('Grace');
  });

  it('keeps the live options merged rather than replaced', () => {
    TestBed.configureTestingModule({
      providers: [provideDocxEditor({ author: 'Ada', title: 'Doc' })],
    });
    const store = TestBed.inject(DocxEditorStore);
    store.configure({ locale: 'es' });
    expect(store.options()).toMatchObject({ author: 'Ada', title: 'Doc', locale: 'es' });
  });

  it('reports host dispatch slots as available only when the host wired them', () => {
    TestBed.configureTestingModule({ providers: [provideDocxEditor({})] });
    const store = TestBed.inject(DocxEditorStore);
    // Open always has a packaged fallback (the file picker), so it is available.
    expect(store.commandState('file.open').enabled).toBe(true);
    // Save has one too (a download).
    expect(store.commandState('file.save').enabled).toBe(true);
    // Export needs a host handler: the engine cannot convert DOCX by itself.
    expect(store.commandState('file.exportPdf').enabled).toBe(false);
    expect(store.commandState('file.exportPdf').disabledReason).not.toBeNull();
  });

  it('enables export once the host supplies a handler', () => {
    TestBed.configureTestingModule({
      providers: [provideDocxEditor({ onExport: () => undefined })],
    });
    const store = TestBed.inject(DocxEditorStore);
    expect(store.commandState('file.exportMarkdown').enabled).toBe(true);
    expect(store.commandState('file.exportPdf').enabled).toBe(true);
  });

  it('delegates open to the host handler', async () => {
    const onOpen = vi.fn();
    TestBed.configureTestingModule({ providers: [provideDocxEditor({ onOpen })] });
    const store = TestBed.inject(DocxEditorStore);
    await store.openFile();
    expect(onOpen).toHaveBeenCalledTimes(1);
  });

  it('answers an honest empty read model before any document is mounted', () => {
    TestBed.configureTestingModule({ providers: [provideDocxEditor({})] });
    const store = TestBed.inject(DocxEditorStore);
    const snapshot = store.snapshot();
    expect(snapshot.isLoading).toBe(true);
    expect(store.editable()).toBe(false);
    expect(store.page().total).toBe(0);
    expect(store.outline()).toEqual([]);
  });

  it('is terminal: destroy releases the instance and ignores later configuration', () => {
    TestBed.configureTestingModule({ providers: [provideDocxEditor({ author: 'Ada' })] });
    const store = TestBed.inject(DocxEditorStore);
    store.destroy();
    expect(store.instance()).toBeNull();
    const optionsAfterDestroy = store.options();
    store.configure({ author: 'Grace' });
    expect(store.options()).toBe(optionsAfterDestroy);
  });

  it('resolves chrome labels through the catalogue', () => {
    TestBed.configureTestingModule({
      providers: [provideDocxEditor({ i18n: { toolbar: { save: 'Guardar' } } as never })],
    });
    const store = TestBed.inject(DocxEditorStore);
    expect(store.t()('toolbar.save')).toBe('Guardar');
    expect(store.t()('toolbar.open')).toBe('Open');
  });
});
