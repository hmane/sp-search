import { StateCreator } from 'zustand';
import {
  ISearchStore,
  IUserSlice,
  ISavedSearch,
  ISearchHistoryEntry,
  ISearchCollection
} from '@interfaces/index';

/**
 * User slice — manages saved searches, search history, and collections state.
 * Components call SearchManagerService directly for CRUD operations,
 * then update the store via these actions.
 */
export const createUserSlice: StateCreator<ISearchStore, [], [], IUserSlice> = (set, get) => ({
  savedSearches: [],
  searchHistory: [],
  collections: [],

  setSavedSearches: (searches: ISavedSearch[]): void => {
    set({ savedSearches: searches });
  },

  setSearchHistory: (history: ISearchHistoryEntry[]): void => {
    set({ searchHistory: history });
  },

  setCollections: (collections: ISearchCollection[]): void => {
    set({ collections });
  },

  addSavedSearch: (search: ISavedSearch): void => {
    const current = get().savedSearches;
    set({ savedSearches: [search, ...current] });
  },

  removeSavedSearch: (id: number): void => {
    const current = get().savedSearches;
    const updated = current.filter((s) => s.id !== id);
    set({ savedSearches: updated });
  },

  addToHistory: (entry: ISearchHistoryEntry): void => {
    const current = get().searchHistory;
    // Deduplicate by queryHash
    const filtered = current.filter((h) => h.queryHash !== entry.queryHash);
    set({ searchHistory: [entry, ...filtered].slice(0, 50) });
  },

  clearSearchHistory: (): void => {
    set({ searchHistory: [] });
  },
});
