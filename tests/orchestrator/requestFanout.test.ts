import { StoreApi } from 'zustand/vanilla';
import { SearchOrchestrator } from '../../src/libraries/spSearchStore/orchestrator/SearchOrchestrator';
import type {
  IFilterConfig,
  IRefiner,
  ISearchDataProvider,
  ISearchQuery,
  ISearchResponse,
  ISearchStore,
} from '../../src/libraries/spSearchStore/interfaces';
import { createMockStore } from '../utils/testHelpers';

function makeResponse(totalCount: number): ISearchResponse {
  return {
    items: [],
    totalCount,
    refiners: [],
    promotedResults: [],
  };
}

function makeFilterConfig(overrides: Partial<IFilterConfig>): IFilterConfig {
  return {
    id: overrides.id || overrides.managedProperty || 'filter',
    displayName: overrides.displayName || overrides.managedProperty || 'Filter',
    managedProperty: overrides.managedProperty || 'FileType',
    filterType: overrides.filterType || 'checkbox',
    operator: overrides.operator || 'OR',
    maxValues: overrides.maxValues || 10,
    defaultExpanded: true,
    showCount: true,
    sortBy: 'count',
    sortDirection: 'desc',
    multiValues: overrides.multiValues !== undefined ? overrides.multiValues : true,
    ...overrides,
  };
}

function refiner(filterName: string, values: Array<{ name: string; value: string; count: number }>): IRefiner {
  return {
    filterName,
    values: values.map((value) => ({
      name: value.name,
      value: value.value,
      count: value.count,
      isSelected: false,
    })),
  };
}

function waitFor(predicate: () => boolean): Promise<void> {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    function check(): void {
      if (predicate()) {
        resolve();
        return;
      }
      if (Date.now() - start > 1000) {
        reject(new Error('Timed out waiting for predicate'));
        return;
      }
      setTimeout(check, 0);
    }
    check();
  });
}

describe('SearchOrchestrator request fan-out', () => {
  let store: StoreApi<ISearchStore>;
  let calls: ISearchQuery[];
  let provider: ISearchDataProvider;

  beforeEach(() => {
    store = createMockStore();
    calls = [];
    provider = {
      id: 'sharepoint-search',
      displayName: 'SharePoint Search',
      supportsRefiners: true,
      supportsCollapsing: true,
      supportsSorting: true,
      execute: jest.fn(async (query: ISearchQuery): Promise<ISearchResponse> => {
        calls.push(query);
        return makeResponse(calls.length * 10);
      }),
    };
    store.getState().registries.dataProviders.register(provider);
  });

  it('does not issue vertical count requests for the active vertical or link-only verticals', async () => {
    store.setState({
      currentVerticalKey: 'all',
      verticals: [
        { key: 'all', label: 'All', sortOrder: 1 },
        { key: 'docs', label: 'Documents', queryTemplate: '{searchTerms} IsDocument:1', sortOrder: 2 },
        { key: 'external', label: 'External', isLink: true, linkUrl: 'https://example.com', sortOrder: 3 },
        { key: 'pages', label: 'Pages', queryTemplate: '{searchTerms} contentclass:STS_ListItem_WebPageLibrary', sortOrder: 4 },
      ],
    });

    const orchestrator = new SearchOrchestrator(store, 0);
    await orchestrator.triggerSearch();

    await waitFor(() => calls.length === 3 && store.getState().verticalCounts.all === 10);

    expect(calls).toHaveLength(3);
    expect(calls[0].pageSize).toBe(25);
    expect(calls.slice(1).map((q) => q.queryTemplate)).toEqual([
      '{searchTerms} IsDocument:1',
      '{searchTerms} contentclass:STS_ListItem_WebPageLibrary',
    ]);
    expect(store.getState().verticalCounts.all).toBe(10);
  });

  it('refreshes an active multi-value refiner without its own filters', async () => {
    store.getState().setAvailableRefiners([
      refiner('FileType', [
        { name: 'xlsx', value: '"xlsx"', count: 462 },
        { name: 'docx', value: '"docx"', count: 462 },
        { name: 'pdf', value: '"pdf"', count: 461 },
      ]),
      refiner('Author', [
        { name: 'Hemant', value: '"Hemant"', count: 200 },
      ]),
    ]);
    store.setState({
      filterConfig: [
        makeFilterConfig({ managedProperty: 'FileType', displayName: 'File type', filterType: 'checkbox' }),
        makeFilterConfig({ managedProperty: 'Author', displayName: 'Author', filterType: 'checkbox' }),
      ],
      activeFilters: [
        { filterName: 'FileType', value: '"xlsx"', displayValue: 'xlsx', operator: 'OR' },
      ],
    });

    calls = [];
    provider.execute = jest.fn(async (query: ISearchQuery): Promise<ISearchResponse> => {
      calls.push(query);
      if (query.pageSize === 0 && query.refiners.length === 1 && query.refiners[0] === 'FileType') {
        return {
          items: [],
          totalCount: 1385,
          refiners: [
            refiner('FileType', [
              { name: 'xlsx', value: '"xlsx"', count: 462 },
              { name: 'docx', value: '"docx"', count: 462 },
              { name: 'pdf', value: '"pdf"', count: 461 },
            ]),
          ],
          promotedResults: [],
        };
      }
      return {
        items: [],
        totalCount: 462,
        refiners: [
          refiner('FileType', [
            { name: 'xlsx', value: '"xlsx"', count: 462 },
          ]),
          refiner('Author', [
            { name: 'Hemant', value: '"Hemant"', count: 17 },
          ]),
        ],
        promotedResults: [],
      };
    });

    const orchestrator = new SearchOrchestrator(store, 0);
    await orchestrator.triggerSearch();

    expect(calls).toHaveLength(2);
    expect(calls[0].filters).toEqual([
      { filterName: 'FileType', value: '"xlsx"', displayValue: 'xlsx', operator: 'OR' },
    ]);
    expect(calls[1].pageSize).toBe(0);
    expect(calls[1].refiners).toEqual(['FileType']);
    expect(calls[1].filters).toEqual([]);
    expect(store.getState().displayRefiners).toEqual([
      refiner('FileType', [
        { name: 'xlsx', value: '"xlsx"', count: 462 },
        { name: 'docx', value: '"docx"', count: 462 },
        { name: 'pdf', value: '"pdf"', count: 461 },
      ]),
      refiner('Author', [
        { name: 'Hemant', value: '"Hemant"', count: 17 },
      ]),
    ]);
  });

  it('issues one self-refiner request per active refiner group, not per selected value', async () => {
    store.getState().setAvailableRefiners([
      refiner('FileType', [
        { name: 'xlsx', value: '"xlsx"', count: 462 },
        { name: 'pdf', value: '"pdf"', count: 461 },
      ]),
    ]);
    store.setState({
      filterConfig: [
        makeFilterConfig({ managedProperty: 'FileType', displayName: 'File type', filterType: 'checkbox' }),
      ],
      activeFilters: [
        { filterName: 'FileType', value: '"xlsx"', displayValue: 'xlsx', operator: 'OR' },
        { filterName: 'FileType', value: '"pdf"', displayValue: 'pdf', operator: 'OR' },
      ],
    });

    calls = [];
    provider.execute = jest.fn(async (query: ISearchQuery): Promise<ISearchResponse> => {
      calls.push(query);
      return {
        items: [],
        totalCount: 10,
        refiners: [
          refiner('FileType', [
            { name: 'xlsx', value: '"xlsx"', count: 462 },
            { name: 'pdf', value: '"pdf"', count: 461 },
          ]),
        ],
        promotedResults: [],
      };
    });

    const orchestrator = new SearchOrchestrator(store, 0);
    await orchestrator.triggerSearch();

    expect(calls).toHaveLength(2);
    expect(calls.filter((query) => query.pageSize === 0 && query.refiners[0] === 'FileType')).toHaveLength(1);
  });

  it('does not self-refresh single-value refiner configurations', async () => {
    store.getState().setAvailableRefiners([
      refiner('FileType', [
        { name: 'xlsx', value: '"xlsx"', count: 462 },
      ]),
    ]);
    store.setState({
      filterConfig: [
        makeFilterConfig({
          managedProperty: 'FileType',
          displayName: 'File type',
          filterType: 'dropdown',
          multiValues: false,
        }),
      ],
      activeFilters: [
        { filterName: 'FileType', value: '"xlsx"', displayValue: 'xlsx', operator: 'OR' },
      ],
    });

    calls = [];
    provider.execute = jest.fn(async (query: ISearchQuery): Promise<ISearchResponse> => {
      calls.push(query);
      return {
        items: [],
        totalCount: 10,
        refiners: [
          refiner('FileType', [
            { name: 'xlsx', value: '"xlsx"', count: 10 },
          ]),
        ],
        promotedResults: [],
      };
    });

    const orchestrator = new SearchOrchestrator(store, 0);
    await orchestrator.triggerSearch();

    expect(calls).toHaveLength(1);
  });
});
