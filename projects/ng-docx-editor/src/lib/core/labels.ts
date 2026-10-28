import { localizeDisabledReason } from '@docx-editor.dev/i18n';
import type { TFunction } from '@docx-editor.dev/i18n';
import type { ChromeSlotId, ToolbarCommandState } from '@docx-editor.dev/core/editor';
import { slotLabelKey } from './chrome-registry';

/** A label resolver, catalogue-backed or host-supplied. */
export type DocxTranslate = (key: string, params?: Record<string, string | number>) => string;

/** One slot's accessible name / tooltip, resolved through the catalogue. */
export function slotLabel(t: DocxTranslate, slot: ChromeSlotId): string {
  return t(slotLabelKey(slot));
}

/** The engine's refusal reason in the reader's language, or the raw reason if unknown. */
export function refusalText(t: DocxTranslate, reason: string | null | undefined): string | null {
  if (!reason) return null;
  return localizeDisabledReason(reason, t as unknown as TFunction) ?? reason;
}

/**
 * A control's `title`: its label, plus the engine's own refusal when the control is
 * disabled — the same rule both adapters follow, so a greyed button always says why.
 */
export function controlTitle(
  t: DocxTranslate,
  slot: ChromeSlotId,
  state: ToolbarCommandState
): string {
  const label = slotLabel(t, slot);
  const reason = state.enabled ? null : refusalText(t, state.disabledReason);
  return reason ? `${label} — ${reason}` : label;
}
