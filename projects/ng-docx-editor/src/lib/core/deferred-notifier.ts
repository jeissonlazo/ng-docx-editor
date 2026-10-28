/**
 * Coalesced deferred notifier for editor state ticks.
 *
 * Ported from `@docx-editor.dev/react` (`packages/react/src/editor/deferred-notifier.ts`,
 * Apache-2.0, © EigenPal). The scheduling policy is part of the engine's contract, not of
 * the React renderer: `Editor.on('change')` fires mid-commit, before the layout scheduler
 * publishes, and a notification storm under sustained typing must yield to a task so the
 * framework's own update guard is never chained past its limit.
 *
 * In Angular the same rules apply for the same reason: a signal written synchronously from
 * the `change` event during a template's command handler would read the not-yet-published
 * layout and cache that stale answer for the rest of the version.
 */

/**
 * How many consecutive notification WAVES may take the microtask path before the next one
 * yields through a task.
 */
const MAX_MICROTASK_WAVES_BEFORE_YIELD = 16;

let waveStreak = 0;
let waveOpen = false;
let streakResetArmed = false;

/**
 * Run `task` in a later event-loop task, through a message rather than a timer: module-level
 * code cannot know which window the editor is in, and a pop-out window's opener throttles
 * timers to one run a second.
 */
let pendingTasks: (() => void)[] = [];
let taskChannel: MessageChannel | null = null;

export function yieldTask(task: () => void): void {
  if (typeof MessageChannel !== 'function') {
    globalThis.setTimeout(task, 0);
    return;
  }
  if (!taskChannel) {
    taskChannel = new MessageChannel();
    taskChannel.port1.onmessage = () => {
      const tasks = pendingTasks;
      pendingTasks = [];
      for (const run of tasks) {
        try {
          run();
        } catch (error) {
          queueMicrotask(() => {
            throw error;
          });
        }
      }
    };
  }
  pendingTasks.push(task);
  if (pendingTasks.length === 1) taskChannel.port2.postMessage(null);
}

/** Decide the scheduling lane for one coalesced notification. */
export function notificationYieldsToTask(): boolean {
  const scheduling = (
    globalThis as typeof globalThis & {
      navigator?: {
        scheduling?: { isInputPending?: (options?: { includeContinuous?: boolean }) => boolean };
      };
    }
  ).navigator?.scheduling;
  if (typeof scheduling?.isInputPending !== 'function') return false;
  if (scheduling.isInputPending({ includeContinuous: true })) return true;
  if (!waveOpen) {
    waveOpen = true;
    waveStreak += 1;
    queueMicrotask(() => {
      waveOpen = false;
    });
    if (!streakResetArmed) {
      streakResetArmed = true;
      yieldTask(() => {
        streakResetArmed = false;
        waveStreak = 0;
      });
    }
  }
  return waveStreak > MAX_MICROTASK_WAVES_BEFORE_YIELD;
}

/** Wrap one notification in the coalescing, deferring lane. */
export function deferredTick(notify: () => void): () => void {
  let scheduled = false;
  return () => {
    if (scheduled) return;
    scheduled = true;
    const flush = () => {
      scheduled = false;
      notify();
    };
    if (notificationYieldsToTask()) yieldTask(flush);
    else queueMicrotask(flush);
  };
}
