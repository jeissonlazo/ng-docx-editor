import { describe, expect, it } from 'vitest';
import { en } from '@docx-editor.dev/i18n';
import { createDocxI18n, DEFAULT_DOCX_I18N } from './i18n';

describe('createDocxI18n', () => {
  it('falls back to bundled English with no override', () => {
    const i18n = createDocxI18n(undefined, undefined, undefined);
    expect(i18n.catalogue).toBe(en);
    expect(i18n.lang).toBe('en');
    expect(i18n.t('toolbar.save')).toBe('Save');
  });

  it('merges a partial catalogue over the inherited one', () => {
    const i18n = createDocxI18n(
      DEFAULT_DOCX_I18N,
      { _lang: 'es', toolbar: { save: 'Guardar' } } as never,
      undefined
    );
    expect(i18n.t('toolbar.save')).toBe('Guardar');
    // Keys the override leaves out fall back rather than showing the key.
    expect(i18n.t('toolbar.open')).toBe('Open');
    expect(i18n.lang).toBe('es');
  });

  it('keeps the inherited language when the override declares none', () => {
    const spanish = createDocxI18n(DEFAULT_DOCX_I18N, { _lang: 'es' } as never, undefined);
    const nested = createDocxI18n(spanish, { toolbar: { save: 'Guardar' } } as never, undefined);
    expect(nested.lang).toBe('es');
  });

  it('keeps the host-page direction unless a provider decides one', () => {
    expect(createDocxI18n(undefined, undefined, undefined).chromeDir).toBeUndefined();
    expect(createDocxI18n(undefined, undefined, 'rtl').chromeDir).toBe('rtl');
    expect(createDocxI18n(undefined, undefined, 'rtl').dir).toBe('rtl');
  });

  it('interpolates parameters', () => {
    const i18n = createDocxI18n(undefined, undefined, undefined);
    expect(i18n.t('viewer.pageIndicator', { current: 2, total: 9 })).toBe('2 of 9');
  });
});
