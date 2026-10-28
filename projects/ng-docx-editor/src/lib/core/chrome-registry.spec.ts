import { describe, expect, it } from 'vitest';
import { CHROME_GROUPS } from '@docx-editor.dev/core/editor';
import type { ChromeSlotId } from '@docx-editor.dev/core/editor';
import { chromeControlForSlot, slotIconPaths, slotInfo, slotLabelKey } from './chrome-registry';

describe('chrome registry projection', () => {
  it('has an entry for every slot the engine registers', () => {
    for (const group of CHROME_GROUPS) {
      for (const control of group.controls) {
        const slot = `${group.id}.${control.id}` as ChromeSlotId;
        expect(slotInfo(slot), slot).toBeDefined();
      }
    }
  });

  it('keeps the registry’s label keys and glyphs', () => {
    expect(slotLabelKey('text.bold')).toBe('formattingBar.boldShortcut');
    expect(slotIconPaths('text.bold')).not.toBeNull();
    expect(slotIconPaths('font.family')).toBeNull();
  });

  it('reports the shape the registry declares', () => {
    expect(slotInfo('text.bold')?.shape).toBe('icon');
    expect(slotInfo('font.family')?.shape).toBe('dropdown');
    expect(slotInfo('font.size')?.shape).toBe('stepper');
    expect(slotInfo('text.color')?.shape).toBe('colorSplit');
  });

  it('flags contextual groups, which the default bar leaves out', () => {
    expect(slotInfo('file.save')?.contextual).toBe(true);
    expect(slotInfo('file.save')?.inDefaultBar).toBe(false);
    expect(slotInfo('text.bold')?.contextual).toBe(false);
    expect(slotInfo('text.bold')?.inDefaultBar).toBe(true);
    expect(slotInfo('image.insert')?.inDefaultBar).toBe(false);
  });

  it('resolves a host back to the engine’s own descriptor', () => {
    const resolved = chromeControlForSlot('text.italic');
    expect(resolved?.group.id).toBe('text');
    expect(resolved?.control.id).toBe('italic');
    expect(resolved?.control.labelKey).toBe('formattingBar.italicShortcut');
  });

  it('answers undefined for a slot that does not exist', () => {
    expect(slotInfo('nope.nothing' as ChromeSlotId)).toBeUndefined();
    expect(chromeControlForSlot('nope.nothing' as ChromeSlotId)).toBeUndefined();
  });
});
