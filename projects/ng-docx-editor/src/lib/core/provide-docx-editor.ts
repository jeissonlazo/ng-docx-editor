import { EnvironmentProviders, makeEnvironmentProviders, Provider } from '@angular/core';
import type { Translations } from '@docx-editor.dev/i18n';
import { DocxEditorStore } from './docx-editor-store';
import type { DocxEditorOptions } from './types';
import { DOCX_EDITOR_DIR, DOCX_EDITOR_I18N, DOCX_EDITOR_SHORTCUTS } from './tokens';

/**
 * Provider-first composition: own one editor at this injector level and hand its instance to
 * every chrome part declared beside or below it.
 *
 * This is the Angular equivalent of React's `DocxEditor.Root` — the lifetime owner — except
 * that the instance is reachable through DI, which is what lets a toolbar, an outline pane
 * and the mount point be siblings rather than ancestors. The editor is created immediately
 * (no container needed) and attached when `<docx-editor-content>` renders.
 *
 * @example
 * ```ts
 * @Component({
 *   providers: [provideDocxEditor({ document: bytes, author: 'Ada' })],
 *   imports: [DocxEditorViewport, DocxEditorContent, DocxToolbar],
 *   template: `
 *     <div class="docx-editor" style="display:flex;flex-direction:column;height:100%">
 *       <docx-toolbar />
 *       <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
 *     </div>
 *   `,
 * })
 * export class MyEditorComponent {}
 * ```
 */
export function provideDocxEditor(options: DocxEditorOptions = {}): Provider[] {
  return [
    {
      // The factory IS the injection context the store is built in, so its own `inject()`
      // calls (DestroyRef, the app-wide i18n tokens) resolve against THIS injector — which is
      // what makes the editor's lifetime the host component's, exactly like a provider-first
      // `DocxEditor.Root` in React.
      provide: DocxEditorStore,
      useFactory: () => {
        const store = new DocxEditorStore();
        store.configure(options);
        return store;
      },
    },
  ];
}

/**
 * App- or feature-level chrome configuration shared by every editor below.
 *
 * @example
 * ```ts
 * bootstrapApplication(App, {
 *   providers: [provideDocxDefaults({ i18n: es, dir: 'ltr' })],
 * });
 * ```
 */
export interface DocxEditorDefaults {
  /** A catalogue every editor merges over bundled English. */
  readonly i18n?: Translations;
  /** Force the chrome writing direction. */
  readonly dir?: 'ltr' | 'rtl';
  /** Bind Ctrl/Cmd+S, O and P inside the editor to the host actions. Default `true`. */
  readonly shortcuts?: boolean;
}

/** Provide app-wide chrome defaults for every editor in the injector subtree. */
export function provideDocxDefaults(defaults: DocxEditorDefaults): EnvironmentProviders {
  return makeEnvironmentProviders([
    ...(defaults.i18n !== undefined
      ? [{ provide: DOCX_EDITOR_I18N, useValue: defaults.i18n }]
      : []),
    ...(defaults.dir !== undefined ? [{ provide: DOCX_EDITOR_DIR, useValue: defaults.dir }] : []),
    ...(defaults.shortcuts !== undefined
      ? [{ provide: DOCX_EDITOR_SHORTCUTS, useValue: defaults.shortcuts }]
      : []),
  ]);
}
