import {
  AUTO_ZOOM_MODE,
  FIT_WIDTH_ZOOM_MODE,
  resolveZoomMode,
  sameZoomMode,
} from '@docx-editor.dev/core/editor';
import type { ZoomMode } from '@docx-editor.dev/core/contracts/editor';

/**
 * The discrete rungs a zoom control steps through.
 *
 * 25% is below the fit floor on purpose: a floored fit sits exactly ON the floor, and a
 * stepper that only offers levels strictly above the current one would go dead in exactly
 * the case the floor exists for — a narrow window with the comments rail open.
 */
export const ZOOM_LEVELS = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3] as const;

/** The next rung above (or below) `zoom`, or `null` at the end of the ladder. */
export function stepZoomLevel(zoom: number, direction: 'in' | 'out'): number | null {
  const epsilon = 0.001;
  return direction === 'in'
    ? (ZOOM_LEVELS.find((level) => level > zoom + epsilon) ?? null)
    : ([...ZOOM_LEVELS].reverse().find((level) => level < zoom - epsilon) ?? null);
}

/** The zoom a Ctrl/Cmd `0`, `+`/`=` or `-` chord asks for, or `null` for any other key. */
export function zoomLevelForShortcut(key: string, zoom: number): number | null {
  if (key === '0') return 1;
  if (key === '+' || key === '=') return stepZoomLevel(zoom, 'in') ?? zoom;
  if (key === '-') return stepZoomLevel(zoom, 'out') ?? zoom;
  return null;
}

/** One row of the zoom dropdown. */
export interface ZoomMenuRow {
  readonly id: string;
  /** The row's label: a percentage for level rows, a catalogue string for fit rows. */
  readonly label: string;
  /** A fixed scale this row sets, if it is a level row. */
  readonly zoom?: number;
  /** A fit mode this row sets, if it is a fit row. */
  readonly mode?: ZoomMode;
}

/** The packaged zoom menu: fixed levels, then the two fit modes. */
export function zoomMenuRows(t: (key: string) => string): readonly ZoomMenuRow[] {
  return [
    ...ZOOM_LEVELS.map((level) => ({
      id: `level:${level}`,
      label: `${Math.round(level * 100)}%`,
      zoom: level,
    })),
    { id: 'fit:fitWidth', label: t('zoom.fitWidth'), mode: FIT_WIDTH_ZOOM_MODE },
    { id: 'auto', label: t('zoom.fitPageWidth'), mode: AUTO_ZOOM_MODE },
  ];
}

/**
 * Whether two declared zoom modes say the same thing, `'auto'` shorthand and an omitted
 * bound included — by VALUE, because a host writing the mode inline hands over a new object
 * on every change.
 */
export function sameDeclaredZoomMode(
  a: ZoomMode | 'auto' | undefined,
  b: ZoomMode | 'auto' | undefined
): boolean {
  if (a === undefined || b === undefined) return a === b;
  const left = resolveZoomMode(a);
  const right = resolveZoomMode(b);
  return left !== null && right !== null && sameZoomMode(left, right);
}

/** Whether the mode currently in force is a fit (so a stepper must not fight it). */
export function isFitZoomMode(mode: ZoomMode | undefined): boolean {
  return mode?.type === 'fit';
}

export { AUTO_ZOOM_MODE, FIT_WIDTH_ZOOM_MODE };
