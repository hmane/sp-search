import type { ISearchHistoryEntry } from '@interfaces/index';
import { getHistoryDisplay } from '@webparts/spSearchManager/components/historyDisplay';

function historyEntry(overrides: Partial<ISearchHistoryEntry>): ISearchHistoryEntry {
  return {
    id: 1,
    queryHash: 'hash',
    queryText: 'Browse all results',
    vertical: 'all',
    searchPageUrl: '',
    searchState: '{}',
    useCount: 1,
    resultCount: 10,
    clickedItems: [],
    searchTimestamp: new Date(2026, 5, 14, 9, 0, 0),
    ...overrides,
  };
}

describe('getHistoryDisplay', () => {
  it('shows filters without the default all vertical prefix', () => {
    const display = getHistoryDisplay(historyEntry({
      queryText: 'all • FileType: pdf',
      searchState: JSON.stringify({
        currentVerticalKey: 'all',
        activeFilters: [
          { filterName: 'FileType', value: 'pdf', displayValue: 'pdf', operator: 'OR' },
        ],
      }),
      useCount: 41,
      resultCount: 460,
    }));

    expect(display.title).toBe('FileType: pdf');
    expect(display.metaParts).toEqual(['Used 41 times', '460 results']);
  });

  it('keeps non-default vertical context visible', () => {
    const display = getHistoryDisplay(historyEntry({
      vertical: 'documents',
      searchState: JSON.stringify({
        queryText: 'budget',
        currentVerticalKey: 'documents',
      }),
      resultCount: 12,
    }));

    expect(display.title).toBe('budget');
    expect(display.metaParts).toEqual(['Scope: documents', '12 results']);
  });

  it('cleans older persisted titles when searchState has no display details', () => {
    const display = getHistoryDisplay(historyEntry({
      queryText: 'all • FileType: xlsx',
      searchState: '{}',
      resultCount: 458,
    }));

    expect(display.title).toBe('FileType: xlsx');
  });
});
