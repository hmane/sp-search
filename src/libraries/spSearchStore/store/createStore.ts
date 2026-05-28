import { createStore as zustandCreateStore, StoreApi } from 'zustand/vanilla';
import { ISearchStore, IRegistryContainer, ISearchScope } from '@interfaces/index';
import {
  createQuerySlice,
  createFilterSlice,
  createResultSlice,
  createVerticalSlice,
  createUISlice,
  createUserSlice
} from './slices';

const defaultScope: ISearchScope = {
  id: 'all',
  label: 'All SharePoint',
};

/**
 * Create a new Zustand store instance that combines all slices.
 * Each search context (identified by searchContextId) gets its own store.
 */
export function createSearchStore(registries: IRegistryContainer): StoreApi<ISearchStore> {
  return zustandCreateStore<ISearchStore>()((...a) => ({
    ...createQuerySlice(...a),
    ...createFilterSlice(...a),
    ...createResultSlice(...a),
    ...createVerticalSlice(...a),
    ...createUISlice(...a),
    ...createUserSlice(...a),

    registries,

    reset: (): void => {
      const [set] = a;
      // AbortController lives on the orchestrator (not the store). Callers
      // who need to cancel an in-flight request must do so via
      // `getOrchestrator(searchContextId).cancelPending()` before resetting.
      set({
        // Query slice defaults
        queryText: '',
        queryTemplate: '{searchTerms}',
        scope: defaultScope,
        suggestions: [],
        isSearching: false,
        // Filter slice defaults
        activeFilters: [],
        availableRefiners: [],
        displayRefiners: [],
        filterConfig: [],
        isRefining: false,
        // Result slice defaults
        items: [],
        totalCount: 0,
        currentPage: 1,
        pageSize: 25,
        sort: undefined,
        sortableProperties: [],
        promotedResults: [],
        isLoading: false,
        error: undefined,
        // Vertical slice defaults
        currentVerticalKey: 'all',
        verticals: [],
        verticalCounts: {},
        // UI slice defaults
        activeLayoutKey: 'list',
        isSearchManagerOpen: false,
        previewPanel: { isOpen: false, item: undefined },
        // User slice defaults
        savedSearches: [],
        searchHistory: [],
        collections: [],
      });
    },

    dispose: (): void => {
      // No-op. The orchestrator owns the in-flight AbortController and is
      // stopped by `disposeStore` before this runs; URL sync cleanup is
      // handled in `disposeStore` ahead of the slice dispose.
    },
  }));
}
