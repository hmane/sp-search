/**
 * T2.D8 — decide whether a URL update should land in the browser's
 * back/forward stack (`pushState`) or quietly replace the current entry
 * (`replaceState`).
 *
 * Audit acceptance signal: "3-query Back/Forward sequence works;
 * filter toggles do not add history". The distinguishing axis is
 * *navigational* vs *incremental*:
 *
 *   - Navigational (push) → a meaningfully different search:
 *     queryText changed, or currentVerticalKey changed (tab switch).
 *   - Incremental (replace) → tweaking the same search:
 *     filter toggles, pagination, sort change, layout switcher.
 *
 * Initial-load also returns false — the first URL write happens during
 * hydration and shouldn't generate a history entry for a search the
 * user never explicitly performed.
 */

export interface IUrlSnapshotForHistory {
  queryText: string;
  currentVerticalKey: string;
}

export function shouldPushHistory(
  previous: IUrlSnapshotForHistory | undefined,
  next: IUrlSnapshotForHistory
): boolean {
  if (!previous) {
    return false;
  }
  if (previous.queryText !== next.queryText) {
    return true;
  }
  if (previous.currentVerticalKey !== next.currentVerticalKey) {
    return true;
  }
  return false;
}
