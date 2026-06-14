import type { ISearchHistoryEntry } from '@interfaces/index';
import {
  formatHistoryTime,
  groupSearchHistoryByDate,
} from '@webparts/spSearchManager/components/historyGrouping';

function historyEntry(id: number, searchTimestamp: Date): ISearchHistoryEntry {
  return {
    id,
    queryHash: 'hash-' + String(id),
    queryText: 'query ' + String(id),
    vertical: 'all',
    searchPageUrl: '',
    searchState: '{}',
    useCount: 1,
    resultCount: 10,
    clickedItems: [],
    searchTimestamp,
  };
}

describe('historyGrouping', () => {
  it('groups history entries by local date while preserving row order', () => {
    const now = new Date(2026, 5, 14, 12, 0, 0);
    const groups = groupSearchHistoryByDate([
      historyEntry(1, new Date(2026, 5, 14, 11, 15, 0)),
      historyEntry(2, new Date(2026, 5, 14, 9, 30, 0)),
      historyEntry(3, new Date(2026, 5, 13, 16, 45, 0)),
      historyEntry(4, new Date(2026, 4, 12, 8, 0, 0)),
    ], now);

    expect(groups.map((group) => group.label)).toEqual([
      'Today',
      'Yesterday',
      'May 12, 2026',
    ]);
    expect(groups.map((group) => group.count)).toEqual([2, 1, 1]);
    expect(groups[0].entries.map((entry) => entry.id)).toEqual([1, 2]);
  });

  it('formats row timestamps as time instead of date-relative labels', () => {
    expect(formatHistoryTime(new Date(2026, 5, 14, 9, 5, 0))).toMatch(/9.*05|09.*05/);
  });
});
