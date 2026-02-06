import { getStore, disposeStore, hasStore } from '../../src/libraries/spSearchStore/store/storeRegistry';

/**
 * Tests for the store registry — getStore, disposeStore, hasStore.
 *
 * The registry is a global singleton Map. Because storeRegistry.ts
 * keeps internal module-level state, we need to clean up between tests
 * by disposing all created stores.
 */
describe('storeRegistry', () => {
  const testIds: string[] = [];

  /** Track IDs created in tests so afterEach can clean them up. */
  function trackId(id: string): string {
    testIds.push(id);
    return id;
  }

  afterEach(() => {
    // Dispose all stores created during this test to avoid leaking state
    for (const id of testIds) {
      if (hasStore(id)) {
        disposeStore(id);
      }
    }
    testIds.length = 0;
  });

  describe('getStore', () => {
    it('should create a store on first call', () => {
      const id = trackId('ctx-create-test');
      const store = getStore(id);

      expect(store).toBeDefined();
      expect(store.getState).toBeDefined();
      expect(store.setState).toBeDefined();
      expect(store.subscribe).toBeDefined();
    });

    it('should return the same instance for the same ID', () => {
      const id = trackId('ctx-same-instance');
      const store1 = getStore(id);
      const store2 = getStore(id);

      expect(store1).toBe(store2);
    });

    it('should return different instances for different IDs', () => {
      const id1 = trackId('ctx-a');
      const id2 = trackId('ctx-b');
      const store1 = getStore(id1);
      const store2 = getStore(id2);

      expect(store1).not.toBe(store2);
    });

    it('should create a store with default state values', () => {
      const id = trackId('ctx-defaults');
      const store = getStore(id);
      const state = store.getState();

      // Query slice defaults
      expect(state.queryText).toBe('');
      expect(state.queryTemplate).toBe('{searchTerms}');
      expect(state.scope).toEqual({ id: 'all', label: 'All SharePoint' });
      expect(state.suggestions).toEqual([]);
      expect(state.isSearching).toBe(false);
      expect(state.abortController).toBeUndefined();

      // Filter slice defaults
      expect(state.activeFilters).toEqual([]);
      expect(state.availableRefiners).toEqual([]);
      expect(state.displayRefiners).toEqual([]);
      expect(state.filterConfig).toEqual([]);
      expect(state.isRefining).toBe(false);

      // Result slice defaults
      expect(state.items).toEqual([]);
      expect(state.totalCount).toBe(0);
      expect(state.currentPage).toBe(1);
      expect(state.pageSize).toBe(25);
      expect(state.sort).toBeUndefined();
      expect(state.promotedResults).toEqual([]);
      expect(state.isLoading).toBe(false);
      expect(state.error).toBeUndefined();

      // Vertical slice defaults
      expect(state.currentVerticalKey).toBe('all');
      expect(state.verticals).toEqual([]);
      expect(state.verticalCounts).toEqual({});

      // UI slice defaults
      expect(state.activeLayoutKey).toBe('list');
      expect(state.isSearchManagerOpen).toBe(false);
      expect(state.previewPanel).toEqual({ isOpen: false, item: undefined });
      expect(state.bulkSelection).toEqual([]);

      // User slice defaults
      expect(state.savedSearches).toEqual([]);
      expect(state.searchHistory).toEqual([]);
      expect(state.collections).toEqual([]);
    });

    it('should create a store with registries', () => {
      const id = trackId('ctx-registries');
      const store = getStore(id);
      const state = store.getState();

      expect(state.registries).toBeDefined();
      expect(state.registries.dataProviders).toBeDefined();
      expect(state.registries.suggestions).toBeDefined();
      expect(state.registries.actions).toBeDefined();
      expect(state.registries.layouts).toBeDefined();
      expect(state.registries.filterTypes).toBeDefined();
    });

    it('should isolate state changes between stores', () => {
      const id1 = trackId('ctx-isolated-1');
      const id2 = trackId('ctx-isolated-2');
      const store1 = getStore(id1);
      const store2 = getStore(id2);

      store1.getState().setQueryText('hello');

      expect(store1.getState().queryText).toBe('hello');
      expect(store2.getState().queryText).toBe('');
    });
  });

  describe('disposeStore', () => {
    it('should remove the store from the registry', () => {
      const id = trackId('ctx-dispose');
      getStore(id);
      expect(hasStore(id)).toBe(true);

      disposeStore(id);
      expect(hasStore(id)).toBe(false);
    });

    it('should abort the in-flight controller on dispose', () => {
      const id = trackId('ctx-dispose-abort');
      const store = getStore(id);

      // Simulate an in-flight search by setting an AbortController
      const controller = new AbortController();
      store.setState({ abortController: controller, isSearching: true });

      expect(controller.signal.aborted).toBe(false);

      disposeStore(id);
      expect(controller.signal.aborted).toBe(true);
    });

    it('should be a no-op if the store does not exist', () => {
      // Should not throw
      expect(() => disposeStore('non-existent-ctx')).not.toThrow();
    });

    it('should allow re-creation after dispose', () => {
      const id = trackId('ctx-recreate');
      const store1 = getStore(id);
      store1.getState().setQueryText('before dispose');

      disposeStore(id);

      const store2 = getStore(id);
      expect(store2).not.toBe(store1);
      expect(store2.getState().queryText).toBe('');
    });
  });

  describe('hasStore', () => {
    it('should return false for an ID that was never created', () => {
      expect(hasStore('never-created')).toBe(false);
    });

    it('should return true for an existing store', () => {
      const id = trackId('ctx-has-true');
      getStore(id);
      expect(hasStore(id)).toBe(true);
    });

    it('should return false after the store is disposed', () => {
      const id = trackId('ctx-has-after-dispose');
      getStore(id);
      disposeStore(id);
      expect(hasStore(id)).toBe(false);
    });
  });
});
