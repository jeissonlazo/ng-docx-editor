import { CHROME_GROUPS } from '@docx-editor.dev/core/editor';
import type { ChromeControl, ChromeGroup, ChromeSlotId } from '@docx-editor.dev/core/editor';

/** The SHAPE a control renders as (mirrors the engine's `ChromeControlShape`). */
export type SlotShape = 'icon' | 'stepper' | 'dropdown' | 'colorSplit';

/** What one chrome slot needs to render, derived from the engine's registry. */
export interface SlotInfo {
  readonly slot: ChromeSlotId;
  readonly groupId: string;
  readonly controlId: string;
  readonly groupLabelKey: string;
  /** i18n key for the accessible name and tooltip. */
  readonly labelKey: string;
  /** i18n key of the placeholder value, for a picker. */
  readonly valueKey?: string;
  readonly shape: SlotShape;
  /** Material Symbols path data, or `null` for a value control. */
  readonly paths: readonly string[] | null;
  /** `false` when the group is contextual (not part of the default bar). */
  readonly inDefaultBar: boolean;
  readonly contextual: boolean;
}

/**
 * The registry through the CONFORMANCE view.
 *
 * `CHROME_GROUPS` is a literal `as const` structure, so the optional members (`shape`,
 * `valueKey`, `contextual`, `defaultToolbar`) are absent from the members that omit them
 * rather than optional on a shared type. This is the same cast the engine's own
 * `defaultChromeGroups()` performs for exactly the same reason.
 */
const GROUPS = CHROME_GROUPS as unknown as readonly ChromeGroup[];

const INFO = new Map<ChromeSlotId, SlotInfo>();

for (const group of GROUPS) {
  for (const control of group.controls) {
    const slot = `${group.id}.${control.id}` as ChromeSlotId;
    INFO.set(slot, {
      slot,
      groupId: group.id,
      controlId: control.id,
      groupLabelKey: group.labelKey,
      labelKey: control.labelKey,
      ...(control.valueKey !== undefined ? { valueKey: control.valueKey } : {}),
      shape: control.shape ?? 'icon',
      paths: control.paths ?? null,
      inDefaultBar: !group.contextual && control.defaultToolbar !== false,
      contextual: group.contextual === true,
    });
  }
}

/** Everything the adapter needs to know about one chrome slot. */
export function slotInfo(slot: ChromeSlotId): SlotInfo | undefined {
  return INFO.get(slot);
}

/** The icon paths for one slot, from the registry. */
export function slotIconPaths(slot: ChromeSlotId): readonly string[] | null {
  return INFO.get(slot)?.paths ?? null;
}

/** The tooltip/accessibility key for one slot, from the registry. */
export function slotLabelKey(slot: ChromeSlotId): string {
  return INFO.get(slot)?.labelKey ?? slot;
}

/** The registry's control descriptor for one slot, for a host building its own chrome. */
export function chromeControlForSlot(
  slot: ChromeSlotId
): { readonly group: ChromeGroup; readonly control: ChromeControl } | undefined {
  const info = INFO.get(slot);
  if (!info) return undefined;
  const group = GROUPS.find((candidate) => candidate.id === info.groupId);
  const control = group?.controls.find((candidate) => candidate.id === info.controlId);
  return group && control ? { group, control } : undefined;
}
