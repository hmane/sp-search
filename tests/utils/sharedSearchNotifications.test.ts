import {
  computeUnacknowledgedShares,
  loadAcknowledgedShareIds,
  saveAcknowledgedShareIds,
  acknowledgeShareIds,
  ACK_STORAGE_KEY_PREFIX,
} from '../../src/libraries/spSearchStore/utils/sharedSearchNotifications';
import type { ISavedSearch } from '../../src/libraries/spSearchStore/interfaces/IStoreSlices';

/**
 * T2.D1 — recipient-side polling helpers for shared-search notifications.
 *
 * `computeUnacknowledgedShares` is the heart of the badge / MessageBar
 * count: filters saved searches to "shared with me, not yet
 * acknowledged" plus a safety guard so the user's own saved searches
 * never appear.
 *
 * `loadAcknowledgedShareIds` / `saveAcknowledgedShareIds` /
 * `acknowledgeShareIds` manage the localStorage-backed acknowledgement
 * set so dismissing the banner in one tab survives a refresh.
 */

function search(id: number, overrides: Partial<ISavedSearch> = {}): ISavedSearch {
  return {
    id,
    title: overrides.title || 'Saved search #' + id,
    queryText: overrides.queryText || 'budget',
    searchState: overrides.searchState || '{}',
    searchUrl: overrides.searchUrl || '',
    entryType: overrides.entryType || 'SavedSearch',
    category: overrides.category || '',
    sharedWith: overrides.sharedWith || [],
    resultCount: overrides.resultCount || 0,
    lastUsed: overrides.lastUsed || new Date('2026-05-13T10:00:00Z'),
    created: overrides.created || new Date('2026-05-12T10:00:00Z'),
    author: overrides.author || { displayText: 'Sender', email: 'sender@contoso.com' },
  };
}

describe('computeUnacknowledgedShares — T2.D1', () => {
  it('returns the empty array when no searches are shared with me', () => {
    const searches = [
      search(1, { entryType: 'SavedSearch' }),
      search(2, { entryType: 'SavedSearch' }),
    ];
    expect(computeUnacknowledgedShares(searches, new Set())).toEqual([]);
  });

  it('returns shared-with-me searches that have not been acknowledged', () => {
    const searches = [
      search(1, { entryType: 'SharedSearch', author: { displayText: 'Alice', email: 'alice@contoso.com' } }),
      search(2, { entryType: 'SharedSearch', author: { displayText: 'Bob', email: 'bob@contoso.com' } }),
    ];
    const result = computeUnacknowledgedShares(searches, new Set());
    expect(result).toHaveLength(2);
    expect(result.map((s) => s.id).sort()).toEqual([1, 2]);
  });

  it('filters out shared searches I have already acknowledged', () => {
    const searches = [
      search(1, { entryType: 'SharedSearch' }),
      search(2, { entryType: 'SharedSearch' }),
      search(3, { entryType: 'SharedSearch' }),
    ];
    const acknowledged = new Set<number>([1, 3]);
    const result = computeUnacknowledgedShares(searches, acknowledged);
    expect(result.map((s) => s.id)).toEqual([2]);
  });

  it('never reports the recipient\'s own saved searches as new shares', () => {
    const searches = [
      // SavedSearch entryType wins — never reported as a new share even if it lives in the same list.
      search(1, { entryType: 'SavedSearch' }),
      search(2, { entryType: 'SharedSearch' }),
    ];
    const result = computeUnacknowledgedShares(searches, new Set());
    expect(result.map((s) => s.id)).toEqual([2]);
  });

  it('returns the empty array on null/undefined input', () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect(computeUnacknowledgedShares(null as any, new Set())).toEqual([]);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect(computeUnacknowledgedShares(undefined as any, new Set())).toEqual([]);
  });
});

describe('acknowledgement storage', () => {
  const userKey = 'jane.smith@contoso.com';
  const storageKey = ACK_STORAGE_KEY_PREFIX + userKey;

  beforeEach(() => {
    try {
      window.localStorage.removeItem(storageKey);
    } catch {
      /* localStorage may be unavailable in jsdom — tests guard below */
    }
  });

  it('loadAcknowledgedShareIds returns an empty Set when no prior state', () => {
    expect(loadAcknowledgedShareIds(userKey)).toEqual(new Set());
  });

  it('saveAcknowledgedShareIds + loadAcknowledgedShareIds round-trip', () => {
    saveAcknowledgedShareIds(userKey, new Set([1, 2, 3]));
    const loaded = loadAcknowledgedShareIds(userKey);
    expect(loaded).toEqual(new Set([1, 2, 3]));
  });

  it('acknowledgeShareIds merges new ids with existing ones', () => {
    saveAcknowledgedShareIds(userKey, new Set([1, 2]));
    const next = acknowledgeShareIds(userKey, [2, 3, 4]);
    expect(next).toEqual(new Set([1, 2, 3, 4]));
    expect(loadAcknowledgedShareIds(userKey)).toEqual(new Set([1, 2, 3, 4]));
  });

  it('acknowledgeShareIds is idempotent when given already-known ids', () => {
    saveAcknowledgedShareIds(userKey, new Set([1, 2, 3]));
    const next = acknowledgeShareIds(userKey, [2]);
    expect(next).toEqual(new Set([1, 2, 3]));
  });

  it('survives corrupt localStorage entries (returns empty set, does not throw)', () => {
    try {
      window.localStorage.setItem(storageKey, 'not-json');
    } catch { /* skip */ }
    expect(loadAcknowledgedShareIds(userKey)).toEqual(new Set());
  });
});
