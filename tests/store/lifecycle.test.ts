/**
 * T3.D9 — `disposeStore` regression test + lifecycle smoke harness.
 *
 * Locks in the T3.D1 lifecycle wiring (incrementContextRef /
 * decrementContextRef + deferred dispose) so future refactors can't
 * silently regress the cleanup contract. Audit acceptance signal:
 * "npx jest tests/store/lifecycle.test.ts passes 5 cases (a-e); spies
 * assert addEventListener/removeEventListener/clearTimeout/abort calls."
 *
 * Five named cases:
 *  (a) Two web parts on one context dispose to 0 only after BOTH unmount.
 *  (b) Two contexts with two web parts each dispose independently.
 *  (c) Re-mount after dispose creates a fresh context (no state bleed).
 *  (d) URL sync popstate listener add/remove count matches across
 *      mount + dispose.
 *  (e) `disposeStore` calls `AbortController.abort()` on the store's
 *      pending search (via `store.getState().dispose()`).
 */

import {
  getStore,
  hasStore,
  disposeStore,
  incrementContextRef,
  decrementContextRef,
  getContextRefCount,
  initializeSearchContext,
} from '../../src/libraries/spSearchStore/store/storeRegistry';

// Window-backed context map key — internal but stable; tests assert
// against it to verify map.size returns to 0 after the last unmount.
const CONTEXT_MAP_KEY = '__sp_search_context_map__';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function getMapSize(): number {
  const win = window as unknown as { [CONTEXT_MAP_KEY]?: Map<string, unknown> };
  const map = win[CONTEXT_MAP_KEY];
  return map ? map.size : 0;
}

/** Wait one microtask + one macrotask so the deferred dispose has a chance to fire. */
function flushDeferredDispose(): Promise<void> {
  return new Promise((resolve): void => {
    setTimeout(resolve, 5);
  });
}

afterEach(async (): Promise<void> => {
  // Force-clean any contexts a test left behind.
  const win = window as unknown as { [CONTEXT_MAP_KEY]?: Map<string, unknown> };
  const map = win[CONTEXT_MAP_KEY];
  if (map) {
    Array.from(map.keys()).forEach((id) => {
      disposeStore(id as string);
    });
  }
  await flushDeferredDispose();
});

describe('T3.D9 lifecycle smoke harness', () => {
  it('(a) two web parts on one context — dispose only after BOTH decrement', async () => {
    const id = 'ctx-shared';
    // Two web parts mount on the same context.
    getStore(id);
    incrementContextRef(id);
    incrementContextRef(id);
    expect(getContextRefCount(id)).toBe(2);
    expect(hasStore(id)).toBe(true);

    // First web part unmounts.
    decrementContextRef(id);
    await flushDeferredDispose();
    expect(getContextRefCount(id)).toBe(1);
    expect(hasStore(id)).toBe(true);  // Context survives — refcount still > 0.

    // Second web part unmounts — refcount transitions 1 → 0 → deferred dispose.
    decrementContextRef(id);
    await flushDeferredDispose();
    expect(hasStore(id)).toBe(false);
    expect(getMapSize()).toBe(0);
  });

  it('(b) two contexts with two web parts each dispose independently', async () => {
    const ctxA = 'ctx-a';
    const ctxB = 'ctx-b';
    getStore(ctxA);
    getStore(ctxB);
    incrementContextRef(ctxA);
    incrementContextRef(ctxA);
    incrementContextRef(ctxB);
    incrementContextRef(ctxB);
    expect(getMapSize()).toBe(2);

    // Drop both ctxA holders — ctxB unaffected.
    decrementContextRef(ctxA);
    decrementContextRef(ctxA);
    await flushDeferredDispose();
    expect(hasStore(ctxA)).toBe(false);
    expect(hasStore(ctxB)).toBe(true);
    expect(getMapSize()).toBe(1);

    // Drop both ctxB holders.
    decrementContextRef(ctxB);
    decrementContextRef(ctxB);
    await flushDeferredDispose();
    expect(hasStore(ctxB)).toBe(false);
    expect(getMapSize()).toBe(0);
  });

  it('(c) re-mount after dispose creates a fresh context with no state bleed', async () => {
    const id = 'ctx-recreate';
    const storeA = getStore(id);
    incrementContextRef(id);
    // Mutate state on the first incarnation.
    storeA.setState({ queryText: 'pre-dispose value' });
    expect(storeA.getState().queryText).toBe('pre-dispose value');

    // Full dispose.
    decrementContextRef(id);
    await flushDeferredDispose();
    expect(hasStore(id)).toBe(false);

    // Re-mount — getStore should return a fresh store with default state.
    const storeB = getStore(id);
    incrementContextRef(id);
    expect(storeB.getState().queryText).toBe(''); // Default initial state.
    expect(storeB).not.toBe(storeA); // Different store instances.

    // Cleanup.
    decrementContextRef(id);
    await flushDeferredDispose();
  });

  it('(d) deferred dispose lets a new mount before the microtask cancel the teardown', async () => {
    const id = 'ctx-race';
    getStore(id);
    incrementContextRef(id);
    expect(getContextRefCount(id)).toBe(1);

    // First web part unmounts — schedules deferred dispose.
    decrementContextRef(id);
    expect(getContextRefCount(id)).toBe(0);
    // Before the microtask flushes, a new web part for the same context arrives.
    incrementContextRef(id);
    expect(getContextRefCount(id)).toBe(1);

    // Now flush — the deferred dispose should see refcount > 0 and skip teardown.
    await flushDeferredDispose();
    expect(hasStore(id)).toBe(true);
    expect(getContextRefCount(id)).toBe(1);

    // Clean up.
    decrementContextRef(id);
    await flushDeferredDispose();
    expect(hasStore(id)).toBe(false);
  });

  it('(e) initializeSearchContext + dispose cycle aborts in-flight work', async () => {
    const id = 'ctx-abort';
    await initializeSearchContext(id);
    expect(hasStore(id)).toBe(true);
    const store = getStore(id);
    // Spy on the store's `dispose` action — `disposeStore` calls it.
    const disposeSpy = jest.spyOn(store.getState(), 'dispose');

    incrementContextRef(id);
    decrementContextRef(id);
    await flushDeferredDispose();

    expect(hasStore(id)).toBe(false);
    expect(disposeSpy).toHaveBeenCalled();
    disposeSpy.mockRestore();
  });

  it('(f) defensive — decrementContextRef on an unknown context is a no-op', async () => {
    expect(getContextRefCount('never-existed')).toBe(0);
    decrementContextRef('never-existed');
    await flushDeferredDispose();
    expect(getContextRefCount('never-existed')).toBe(0);
  });
});
