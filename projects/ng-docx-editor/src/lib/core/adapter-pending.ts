import type { ChromeSlotId } from '@docx-editor.dev/core/editor';

/**
 * Slots this adapter does not render yet because they need chrome it does not ship: a dialog,
 * a popover, or a rail.
 *
 * They are OMITTED rather than rendered dead. A disabled control whose reason is "render
 * DocxEditor.PageSetupDialog" tells a reader nothing they can act on, and the engine's own
 * reason for a slot that would work is then indistinguishable from a feature that is simply
 * missing. A host that composes the slot itself (its own page-setup form, its own bookmark
 * picker) renders it and drives it through `DocxEditorStore.run`.
 *
 * Roadmap: page setup, paragraph, watermark and bookmark dialogs; the image properties, wrap
 * and alt-text popovers; the review pane and comment drafts; the table chrome.
 */
export const ADAPTER_PENDING_SLOTS: ReadonlySet<ChromeSlotId> = new Set<ChromeSlotId>([
  'file.pageSetup',
  'paragraph.dialog',
  'insert.bookmark',
  'insert.goToBookmark',
  'insert.watermark',
  'review.revisionMarkup',
  'review.authors',
  'image.properties',
  'image.wrap',
  'image.altText',
]);

/** Whether the adapter renders this slot itself, so chrome should skip it. */
export function isAdapterPending(slot: ChromeSlotId): boolean {
  return ADAPTER_PENDING_SLOTS.has(slot);
}
