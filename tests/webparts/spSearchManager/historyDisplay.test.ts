import type { IFilterConfig, ISearchHistoryEntry } from '@interfaces/index';
import { buildFilterAliasMap, getHistoryDisplay } from '@webparts/spSearchManager/components/historyDisplay';

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

  it('uses the refiner alias instead of the managed-property name when provided', () => {
    const aliases = buildFilterAliasMap([
      { managedProperty: 'RefinableString06', displayName: 'Status' } as IFilterConfig,
    ]);
    const display = getHistoryDisplay(historyEntry({
      searchState: JSON.stringify({
        activeFilters: [
          { filterName: 'RefinableString06', value: '1', displayValue: 'Yes', operator: 'OR' },
        ],
      }),
    }), aliases);

    expect(display.title).toBe('Status: Yes');
  });

  it('decodes a SharePoint hex refinement token to its label', () => {
    const aliases = buildFilterAliasMap([
      { managedProperty: 'Accounts', displayName: 'Accounts' } as IFilterConfig,
    ]);
    // ǂǂ (U+01C2 ×2) + hex "31393534" → "1954". Build from char codes so the
    // test doesn't depend on source-file encoding.
    const token = String.fromCharCode(0x01C2, 0x01C2) + '31393534';
    const display = getHistoryDisplay(historyEntry({
      searchState: JSON.stringify({
        activeFilters: [
          { filterName: 'Accounts', value: token, operator: 'OR' },
        ],
      }),
    }), aliases);

    expect(display.title).toBe('Accounts: 1954');
  });

  it('falls back to the managed-property name when no alias is configured', () => {
    const display = getHistoryDisplay(historyEntry({
      searchState: JSON.stringify({
        activeFilters: [
          { filterName: 'RefinableString06', value: '1', displayValue: 'Yes', operator: 'OR' },
        ],
      }),
    }), buildFilterAliasMap([]));

    expect(display.title).toBe('RefinableString06: Yes');
  });
});

describe('buildFilterAliasMap', () => {
  it('maps managed property (lowercased) to display name, skipping blanks', () => {
    const map = buildFilterAliasMap([
      { managedProperty: 'RefinableString06', displayName: 'Status' } as IFilterConfig,
      { managedProperty: 'Author', displayName: '' } as IFilterConfig,
      { managedProperty: '', displayName: 'Ignored' } as IFilterConfig,
    ]);
    expect(map['refinablestring06']).toBe('Status');
    expect(map.author).toBeUndefined();
    expect(Object.keys(map)).toHaveLength(1);
  });

  it('returns an empty map for undefined config', () => {
    expect(buildFilterAliasMap(undefined)).toEqual({});
  });
});
