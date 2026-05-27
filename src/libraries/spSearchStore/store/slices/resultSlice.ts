import { StateCreator } from 'zustand';
import { ISearchStore, IResultSlice, ISearchResult, ISortField, IPromotedResultItem } from '@interfaces/index';

export const createResultSlice: StateCreator<ISearchStore, [], [], IResultSlice> = (set) => ({
  items: [],
  totalCount: 0,
  currentPage: 1,
  pageSize: 25,
  sort: undefined,
  sortableProperties: [],
  promotedResults: [],
  querySuggestion: undefined,
  // T1.D4 — initial `false` so a freshly-mounted Results web part with no
  // auto-search renders the idle EmptyState ("Enter a search term to get
  // started.") instead of a misleading 5-row shimmer. The orchestrator
  // calls `setLoading(true)` before each search and `setLoading(false)`
  // after, so the shimmer still appears during real in-flight loads.
  isLoading: false,
  hasSearched: false,
  error: undefined,
  resultSourceId: '',
  enableQueryRules: true,
  trimDuplicates: true,
  refinementFilters: '',
  collapseSpecification: '',
  showPaging: true,
  pageRange: 5,
  selectedProperties: '',

  setResults: (items: ISearchResult[], total: number): void => {
    set({ items, totalCount: total, error: undefined, hasSearched: true });
  },

  setPage: (page: number): void => {
    set({ currentPage: page });
  },

  setSort: (sort: ISortField): void => {
    set({ sort, currentPage: 1 });
  },

  setPromotedResults: (results: IPromotedResultItem[]): void => {
    set({ promotedResults: results });
  },

  setQuerySuggestion: (suggestion: string | undefined): void => {
    set({ querySuggestion: suggestion });
  },

  setLoading: (isLoading: boolean): void => {
    set({ isLoading });
  },

  setError: (error: string | undefined): void => {
    if (error !== undefined) {
      set({ error, isLoading: false, hasSearched: true });
    } else {
      set({ error: undefined });
    }
  },
});
