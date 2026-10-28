import { Component, signal } from '@angular/core';
// Imports come from the PUBLISHED package specifier, not from a source path: this project is
// what proves the built artifact (FESM + declarations + exports map) is consumable.
import { DocxEditorComponent } from 'ng-docx-editor';
import type { DocxFonts } from 'ng-docx-editor';
import { defaultFonts } from '@docx-editor.dev/fonts';

/**
 * The smallest real consumer of `ng-docx-editor`: one tag, one document, the packaged fonts.
 * `scripts/check-consumer-install.mjs` builds it from the packed tarball and renders it in a
 * browser, so a broken `exports` map, a missing `styles.css` or a declaration that stops
 * compiling fails the check instead of shipping.
 */
@Component({
  selector: 'smoke-root',
  imports: [DocxEditorComponent],
  template: `
    <docx-editor
      [document]="bytes()"
      [fonts]="fonts()"
      title="smoke.docx"
      [navigation]="false"
    />
  `,
  styles: [
    `
      :host {
        display: block;
        height: 100vh;
      }
    `,
  ],
})
export class SmokeApp {
  readonly bytes = signal<Uint8Array | 'blank' | undefined>(undefined);
  readonly fonts = signal<DocxFonts | undefined>(undefined);

  constructor() {
    void this.load();
  }

  private async load(): Promise<void> {
    const [bytes, fonts] = await Promise.all([
      fetch('sample.docx')
        .then((response) => response.arrayBuffer())
        .then((buffer) => new Uint8Array(buffer)),
      defaultFonts().catch(() => undefined),
    ]);
    this.fonts.set(fonts);
    this.bytes.set(bytes);
  }
}
