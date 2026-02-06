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
      set({
        // Query slice defaults
        queryText: '',
        queryTemplate: '{searchTerms}',
        scope: defaultScope,
        suggestions: [],
        isSearching: false,
        abortController: undefined,
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
        bulkSelection: [],
        // User slice defaults
        savedSearches: [],
        searchHistory: [],
        collections: [],
      });
    },

    dispose: (): void => {
      const [, get] = a;
      // Abort any in-flight search
      const controller = get().abortController;
      if (controller) {
        controller.abort();
      }
      // URL sync cleanup will be added in Step 1.4
    },
  }));
}
