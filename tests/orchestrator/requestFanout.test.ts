import { StoreApi } from 'zustand/vanilla';
import { SearchOrchestrator } from '../../src/libraries/spSearchStore/orchestrator/SearchOrchestrator';
import type {
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
});
