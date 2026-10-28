import { describe, expect, it, vi } from 'vitest';
import { deferredTick } from './deferred-notifier';

describe('deferredTick', () => {
  it('coalesces a burst of notifications into one call', async () => {
    const notify = vi.fn();
    const tick = deferredTick(notify);
    tick();
    tick();
    tick();
    expect(notify).not.toHaveBeenCalled();
    await Promise.resolve();
    expect(notify).toHaveBeenCalledTimes(1);
  });

  it('re-arms after the flush', async () => {
    const notify = vi.fn();
    const tick = deferredTick(notify);
    tick();
    await Promise.resolve();
    tick();
    await Promise.resolve();
    expect(notify).toHaveBeenCalledTimes(2);
  });

  it('defers to a later task, not a microtask, when input is pending', async () => {
    const notify = vi.fn();
    const scheduling = { isInputPending: () => true };
    vi.stubGlobal('navigator', { scheduling });
    try {
      const tick = deferredTick(notify);
      tick();
      await Promise.resolve();
      // The microtask checkpoint passed and the task has not run yet.
      expect(notify).not.toHaveBeenCalled();
      await new Promise((resolve) => setTimeout(resolve, 10));
      expect(notify).toHaveBeenCalledTimes(1);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
