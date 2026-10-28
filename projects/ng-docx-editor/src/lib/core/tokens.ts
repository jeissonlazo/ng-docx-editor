import { InjectionToken } from '@angular/core';
import type { Translations } from '@docx-editor.dev/i18n';

/**
 * App-wide chrome catalogue, overriding the bundled English one.
 *
 * Provide a `Translations` partial (see `@docx-editor.dev/i18n`) and every editor in the
 * subtree resolves its labels through it. An individual `<docx-editor [i18n]="…">` still
 * wins for that editor only.
 *
 * @example
 * ```ts
 * import { es } from '@docx-editor.dev/i18n';
 *
 * bootstrapApplication(App, {
 *   providers: [{ provide: DOCX_EDITOR_I18N, useValue: es }],
 * });
 * ```
 */
export const DOCX_EDITOR_I18N = new InjectionToken<Translations>('DOCX_EDITOR_I18N');

/**
 * The chrome writing direction, when the app wants to force one.
 *
 * Omitted, each editor derives it from its catalogue (`_dir`, then an RTL `_lang`) and
 * otherwise inherits the host page's direction. Forcing `'ltr'` keeps a left-to-right chrome
 * on a right-to-left page; `'rtl'` mirrors the packaged chrome.
 */
export const DOCX_EDITOR_DIR = new InjectionToken<'ltr' | 'rtl'>('DOCX_EDITOR_DIR');

/**
 * Default shortcut handling for Ctrl/Cmd+S, Ctrl/Cmd+O and Ctrl/Cmd+P inside the editor.
 *
 * `true` (the default) binds them to the host actions (`onSave`, `onOpen`, print). `false`
 * leaves every chord to the browser and to the host application.
 */
export const DOCX_EDITOR_SHORTCUTS = new InjectionToken<boolean>('DOCX_EDITOR_SHORTCUTS');
