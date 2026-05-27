import type { ISavedSearch } from '@interfaces/index';

/**
 * T2.D1 — recipient-side notification helpers for shared-search alerts.
 *
 * The Search Manager polls `loadSavedSearches` on a timer; this util
 * does the diffing + acknowledgement bookkeeping so the panel can:
 *   - render a badge showing how many shared searches are unread
 *   - render a MessageBar listing them with a "Got it" dismiss
 *
 * Acknowledgement is stored per-user in localStorage so dismissing
 * the banner survives a tab refresh. Cross-device sync is out of
 * scope; if it matters later, swap the localStorage backend for a
 * SP list field.
 *
 * The "shared with me" predicate uses `entryType === 'SharedSearch'`
 * which `SearchManagerService.loadSavedSearches` already sets when
 * the SharedWith assignment is present — see service line 116.
 */

export const ACK_STORAGE_KEY_PREFIX = 'sp-search-shared-ack-';

/**
 * Filter a savedSearches array to entries shared with the current user
 * that they haven't acknowledged yet. `acknowledgedIds` is the Set
 * loaded by `loadAcknowledgedShareIds` for the current user.
 */
export function computeUnacknowledgedShares(
  savedSearches: ISavedSearch[] | undefined,
  acknowledgedIds: Set<number>
): ISavedSearch[] {
  if (!savedSearches) {
    return [];
  }
  const result: ISavedSearch[] = [];
  for (let i: number = 0; i < savedSearches.length; i++) {
    const item = savedSearches[i];
    if (item.entryType !== 'SharedSearch') {
      continue;
    }
    if (acknowledgedIds.has(item.id)) {
      continue;
    }
    result.push(item);
  }
  return result;
}

function buildStorageKey(userKey: string): string {
  return ACK_STORAGE_KEY_PREFIX + (userKey || 'anonymous');
}

/**
 * Read the set of share IDs the current user has acknowledged from
 * localStorage. Returns an empty Set on any error (missing key,
 * corrupt JSON, localStorage unavailable).
 */
export function loadAcknowledgedShareIds(userKey: string): Set<number> {
  if (typeof window === 'undefined' || !window.localStorage) {
    return new Set();
  }
  try {
    const raw = window.localStorage.getItem(buildStorageKey(userKey));
    if (!raw) {
      return new Set();
    }
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) {
      return new Set();
    }
    const result = new Set<number>();
    for (let i: number = 0; i < parsed.length; i++) {
      if (typeof parsed[i] === 'number') {
        result.add(parsed[i] as number);
      }
    }
    return result;
  } catch {
    return new Set();
  }
}

export function saveAcknowledgedShareIds(userKey: string, ids: Set<number>): void {
  if (typeof window === 'undefined' || !window.localStorage) {
    return;
  }
  try {
    const arr = Array.from(ids);
    window.localStorage.setItem(buildStorageKey(userKey), JSON.stringify(arr));
  } catch {
    // localStorage full / disabled — silent fail so the panel stays usable.
  }
}

/**
 * Merge `newIds` into the user's acknowledged set and persist. Returns
 * the merged Set so callers can update their React state in the same
 * tick.
 */
export function acknowledgeShareIds(userKey: string, newIds: number[]): Set<number> {
  const existing = loadAcknowledgedShareIds(userKey);
  for (let i: number = 0; i < newIds.length; i++) {
    existing.add(newIds[i]);
  }
  saveAcknowledgedShareIds(userKey, existing);
  return existing;
}
