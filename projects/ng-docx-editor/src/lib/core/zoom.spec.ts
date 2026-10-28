import { describe, expect, it } from 'vitest';
import { AUTO_ZOOM_MODE, FIT_WIDTH_ZOOM_MODE } from '@docx-editor.dev/core/editor';
import {
  sameDeclaredZoomMode,
  stepZoomLevel,
  zoomLevelForShortcut,
  zoomMenuRows,
  ZOOM_LEVELS,
} from './zoom';

describe('zoom helpers', () => {
  it('steps to the next rung in each direction, and stops at the ends', () => {
    expect(stepZoomLevel(1, 'in')).toBe(1.25);
    expect(stepZoomLevel(1, 'out')).toBe(0.75);
    expect(stepZoomLevel(3, 'in')).toBeNull();
    expect(stepZoomLevel(0.25, 'out')).toBeNull();
  });

  it('offers 25% so a floored fit can still step down', () => {
    expect(ZOOM_LEVELS[0]).toBe(0.25);
    expect(stepZoomLevel(0.5, 'out')).toBe(0.25);
  });

  it('maps the platform chords', () => {
    expect(zoomLevelForShortcut('0', 1.5)).toBe(1);
    expect(zoomLevelForShortcut('=', 1)).toBe(1.25);
    expect(zoomLevelForShortcut('+', 1)).toBe(1.25);
    expect(zoomLevelForShortcut('-', 1)).toBe(0.75);
    expect(zoomLevelForShortcut('-', 0.25)).toBe(0.25);
    expect(zoomLevelForShortcut('q', 1)).toBeNull();
  });

  it('lists levels then the two fit modes, localized', () => {
    const rows = zoomMenuRows((key) => (key === 'zoom.fitWidth' ? 'Fit width' : key));
    expect(rows.length).toBe(ZOOM_LEVELS.length + 2);
    expect(rows.at(-2)?.label).toBe('Fit width');
    expect(rows.at(-2)?.mode).toBe(FIT_WIDTH_ZOOM_MODE);
    expect(rows.at(-1)?.mode).toBe(AUTO_ZOOM_MODE);
  });

  it('compares declared modes by value, shorthand included', () => {
    expect(sameDeclaredZoomMode('auto', AUTO_ZOOM_MODE)).toBe(true);
    expect(sameDeclaredZoomMode(FIT_WIDTH_ZOOM_MODE, { type: 'fixed' })).toBe(false);
    expect(sameDeclaredZoomMode(undefined, undefined)).toBe(true);
    expect(sameDeclaredZoomMode(undefined, { type: 'fixed' })).toBe(false);
  });
});
