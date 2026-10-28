import { catalogDirection, chromeDirection, createT, deepMerge, en } from '@docx-editor.dev/i18n';
import type { LocaleStrings, TFunction, Translations } from '@docx-editor.dev/i18n';

/** The label resolver every packaged control takes. */
export type DocxTranslate = (key: string, params?: Record<string, string | number>) => string;

/** Everything a chrome part needs to render its labels, in one immutable value. */
export interface DocxI18n {
  /** The merged catalogue the resolver reads. */
  readonly catalogue: LocaleStrings;
  /** BCP-47 tag of the catalogue; formats numbers the way its strings read. */
  readonly lang: string;
  /** The direction the chrome should paint in. */
  readonly dir: 'ltr' | 'rtl';
  /** The direction a provider DECIDED, or `undefined` when the chrome inherits the page. */
  readonly chromeDir: 'ltr' | 'rtl' | undefined;
  /** The catalogue-backed resolver. */
  readonly t: TFunction;
}

/**
 * Build one editor's i18n state by merging `i18n` onto the catalogue it inherits (bundled
 * English at the top of the tree), exactly as the React and Vue adapters do.
 */
export function createDocxI18n(
  inherited: DocxI18n | undefined,
  i18n: Translations | undefined,
  dir: 'ltr' | 'rtl' | undefined
): DocxI18n {
  const baseCatalogue = inherited?.catalogue ?? en;
  const lang = typeof i18n?._lang === 'string' ? i18n._lang : (inherited?.lang ?? 'en');
  const chromeDir = chromeDirection(dir, i18n, inherited?.chromeDir);
  const direction = chromeDir ?? catalogDirection(i18n) ?? inherited?.dir ?? 'ltr';
  const catalogue =
    i18n === undefined
      ? baseCatalogue
      : (deepMerge(baseCatalogue as Record<string, unknown>, i18n as Record<string, unknown>) as
          | LocaleStrings
          | undefined) ?? baseCatalogue;
  return {
    catalogue,
    lang,
    dir: direction,
    chromeDir,
    t: createT(catalogue, lang),
  };
}

/** The default i18n state: bundled English, host-page direction. */
export const DEFAULT_DOCX_I18N: DocxI18n = {
  catalogue: en,
  lang: 'en',
  dir: 'ltr',
  chromeDir: undefined,
  t: createT(en, 'en'),
};
