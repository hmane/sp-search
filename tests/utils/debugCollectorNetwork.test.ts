import { DebugCollector } from '../../src/libraries/spSearchStore/debug/DebugCollector';
import { REDACTED_PLACEHOLDER } from '../../src/libraries/spSearchStore/utils/spLog';

/**
 * T5.D2 — DebugCollector Network buffer: queryText must always be
 * stored as `[redacted]` regardless of what the caller passes, and the
 * ring buffer must cap at 50 rows (audit acceptance signal).
 *
 * The collector reads `window.location.search` once to decide whether
 * it's active. Tests set `?debug=1` before the first call so the
 * `isActive()` check passes; the window-backed state is reset between
 * specs via `__sp_search_debug_collector__`.
 */

function activateDebugAndReset(): void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const win = window as any;
  // Force the collector's `isActive` to re-read by clearing the cached
  // module-level flag (the singleton lives on window).
  delete win.__sp_search_debug_collector__;
  Object.defineProperty(window, 'location', {
    value: { search: '?debug=1', hostname: 'localhost' },
    writable: true,
  });
  // The module caches `_active` per-bundle, so import again to refresh.
  jest.resetModules();
}

describe('DebugCollector.logNetworkEvent — T5.D2', () => {
  beforeEach(() => {
    activateDebugAndReset();
  });

  it('stores queryText as [redacted] regardless of the omitted field', () => {
    // The shape OMITS queryText (the API doesn't accept it) — the collector
    // always inserts the redacted sentinel.
    const Module = jest.requireActual('../../src/libraries/spSearchStore/debug/DebugCollector');
    const collector = Module.DebugCollector;
    collector.logNetworkEvent({
      providerId: 'sp-search',
      kind: 'search',
      status: 'ok',
      durationMs: 42,
      queryTemplate: '{searchTerms}',
      currentPage: 1,
      pageSize: 10,
      totalCount: 7,
      itemCount: 7,
      errorMessage: undefined,
      verticalKey: undefined,
    });
    const events = collector.getNetworkEvents();
    expect(events).toHaveLength(1);
    expect(events[0].queryText).toBe(REDACTED_PLACEHOLDER);
  });

  it('caps the ring buffer at 50 rows — older entries fall off', () => {
    const Module = jest.requireActual('../../src/libraries/spSearchStore/debug/DebugCollector');
    const collector = Module.DebugCollector;
    for (let i: number = 0; i < 60; i++) {
      collector.logNetworkEvent({
        providerId: 'sp-search',
        kind: 'search',
        status: 'ok',
        durationMs: 10 + i,
        queryTemplate: '{searchTerms}',
        currentPage: 1,
        pageSize: 10,
        totalCount: i,
        itemCount: i,
        errorMessage: undefined,
        verticalKey: undefined,
      });
    }
    const events = collector.getNetworkEvents();
    expect(events.length).toBe(50);
    // unshift puts newest at index 0 — so totalCount on the freshest row is 59.
    expect(events[0].totalCount).toBe(59);
    // Oldest surviving row is the 50th-from-newest → totalCount 10.
    expect(events[events.length - 1].totalCount).toBe(10);
  });

  it('preserves status and verticalKey for vertical-count traffic', () => {
    const Module = jest.requireActual('../../src/libraries/spSearchStore/debug/DebugCollector');
    const collector = Module.DebugCollector;
    collector.logNetworkEvent({
      providerId: 'sp-search',
      kind: 'verticalCount',
      status: 'ok',
      durationMs: 88,
      queryTemplate: '{searchTerms}',
      currentPage: 1,
      pageSize: 0,
      totalCount: 17,
      itemCount: 0,
      errorMessage: undefined,
      verticalKey: 'documents',
    });
    const events = collector.getNetworkEvents();
    expect(events[0].kind).toBe('verticalCount');
    expect(events[0].verticalKey).toBe('documents');
  });

  it('records the error path with an aborted/error status', () => {
    const Module = jest.requireActual('../../src/libraries/spSearchStore/debug/DebugCollector');
    const collector = Module.DebugCollector;
    collector.logNetworkEvent({
      providerId: 'sp-search',
      kind: 'search',
      status: 'error',
      durationMs: 250,
      queryTemplate: '{searchTerms}',
      currentPage: 1,
      pageSize: 10,
      totalCount: undefined,
      itemCount: undefined,
      errorMessage: 'HTTP 503',
      verticalKey: undefined,
    });
    const events = collector.getNetworkEvents();
    expect(events[0].status).toBe('error');
    expect(events[0].errorMessage).toBe('HTTP 503');
  });

  it('assigns monotonically increasing ids', () => {
    const Module = jest.requireActual('../../src/libraries/spSearchStore/debug/DebugCollector');
    const collector = Module.DebugCollector;
    for (let i: number = 0; i < 3; i++) {
      collector.logNetworkEvent({
        providerId: 'sp-search',
        kind: 'search',
        status: 'ok',
        durationMs: 10,
        queryTemplate: '{searchTerms}',
        currentPage: 1,
        pageSize: 10,
        totalCount: 0,
        itemCount: 0,
        errorMessage: undefined,
        verticalKey: undefined,
      });
    }
    const events = collector.getNetworkEvents();
    // unshift order: newest first; ids increment with each insert.
    expect(events[0].id).toBeGreaterThan(events[1].id);
    expect(events[1].id).toBeGreaterThan(events[2].id);
  });
});
