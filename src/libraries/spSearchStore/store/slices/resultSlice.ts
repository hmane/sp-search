import { StateCreator } from 'zustand';
import { ISearchStore, IResultSlice, ISearchResult, ISortField, IPromotedResultItem } from '@interfaces/index';

export const createResultSlice: StateCreator<ISearchStore, [], [], IResultSlice> = (set) => ({
  items: [],
  totalCount: 0,
  currentPage: 1,
  pageSize: 25,
  sort: undefined,
  promotedResults: [],
  isLoading: false,
  error: undefined,

  setResults: (items: ISearchResult[], total: number): void => {
    set({ items, totalCount: total, isLoading: false, error: undefined });
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

  setLoading: (isLoading: boolean): void => {
    set({ isLoading });
  },

  setError: (error: string | undefined): void => {
    set({ error, isLoading: false });
  },
});
