import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '../../../src/libraries/spSearchStore/interfaces';
import { createMockStore, createMockSavedSearch, createMockHistoryEntry } from '../../utils/testHelpers';

describe('userSlice', () => {
  let store: StoreApi<ISearchStore>;

  beforeEach(() => {
    store = createMockStore();
  });

  describe('initial state', () => {
    it('should have empty savedSearches', () => {
      expect(store.getState().savedSearches).toEqual([]);
    });

    it('should have empty searchHistory', () => {
      expect(store.getState().searchHistory).toEqual([]);
    });

    it('should have empty collections', () => {
      expect(store.getState().collections).toEqual([]);
    });
  });

  describe('saveSearch', () => {
    it('should add a saved search to the list', async () => {
      const search = createMockSavedSearch({ id: 1, title: 'Budget Reports' });
      await store.getState().saveSearch(search);

      expect(store.getState().savedSearches).toHaveLength(1);
      expect(store.getState().savedSearches[0].title).toBe('Budget Reports');
    });

    it('should append multiple saved searches', async () => {
      await store.getState().saveSearch(createMockSavedSearch({ id: 1, title: 'First' }));
      await store.getState().saveSearch(createMockSavedSearch({ id: 2, title: 'Second' }));

      expect(store.getState().savedSearches).toHaveLength(2);
    });
  });

  describe('addToHistory', () => {
    it('should prepend a history entry', () => {
      const entry = createMockHistoryEntry({ id: 1, queryText: 'first search' });
      store.getState().addToHistory(entry);

      expect(store.getState().searchHistory).toHaveLength(1);
      expect(store.getState().searchHistory[0].queryText).toBe('first search');
    });

    it('should prepend new entries (newest first)', () => {
      store.getState().addToHistory(createMockHistoryEntry({ id: 1, queryText: 'first' }));
      store.getState().addToHistory(createMockHistoryEntry({ id: 2, queryText: 'second' }));
      store.getState().addToHistory(createMockHistoryEntry({ id: 3, queryText: 'third' }));

      expect(store.getState().searchHistory).toHaveLength(3);
      expect(store.getState().searchHistory[0].queryText).toBe('third');
      expect(store.getState().searchHistory[1].queryText).toBe('second');
      expect(store.getState().searchHistory[2].queryText).toBe('first');
    });
  });

  describe('loadHistory', () => {
    it('should be callable without error (Phase 3 stub)', async () => {
      // loadHistory is a Phase 3 stub — ensure it doesn't throw
      await expect(store.getState().loadHistory()).resolves.toBeUndefined();
    });
  });
});
