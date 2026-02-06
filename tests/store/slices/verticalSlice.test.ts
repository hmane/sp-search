import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '../../../src/libraries/spSearchStore/interfaces';
import { createMockStore } from '../../utils/testHelpers';

describe('verticalSlice', () => {
  let store: StoreApi<ISearchStore>;

  beforeEach(() => {
    store = createMockStore();
  });

  describe('initial state', () => {
    it('should default to "all" vertical', () => {
      expect(store.getState().currentVerticalKey).toBe('all');
    });

    it('should have empty verticals array', () => {
      expect(store.getState().verticals).toEqual([]);
    });

    it('should have empty verticalCounts', () => {
      expect(store.getState().verticalCounts).toEqual({});
    });
  });

  describe('setVertical', () => {
    it('should update currentVerticalKey', () => {
      store.getState().setVertical('documents');
      expect(store.getState().currentVerticalKey).toBe('documents');
    });

    it('should handle various vertical keys', () => {
      const keys = ['all', 'documents', 'people', 'sites', 'news', 'videos'];
      for (const key of keys) {
        store.getState().setVertical(key);
        expect(store.getState().currentVerticalKey).toBe(key);
      }
    });

    it('should replace previous vertical', () => {
      store.getState().setVertical('documents');
      store.getState().setVertical('people');
      expect(store.getState().currentVerticalKey).toBe('people');
    });
  });

  describe('setVerticalCounts', () => {
    it('should update verticalCounts', () => {
      const counts = { all: 100, documents: 45, people: 12, sites: 5 };
      store.getState().setVerticalCounts(counts);
      expect(store.getState().verticalCounts).toEqual(counts);
    });

    it('should replace previous counts', () => {
      store.getState().setVerticalCounts({ all: 100 });
      store.getState().setVerticalCounts({ all: 50, documents: 30 });

      expect(store.getState().verticalCounts).toEqual({ all: 50, documents: 30 });
    });

    it('should handle empty counts', () => {
      store.getState().setVerticalCounts({ all: 100 });
      store.getState().setVerticalCounts({});
      expect(store.getState().verticalCounts).toEqual({});
    });

    it('should handle zero counts', () => {
      store.getState().setVerticalCounts({ all: 0, documents: 0 });
      expect(store.getState().verticalCounts.all).toBe(0);
      expect(store.getState().verticalCounts.documents).toBe(0);
    });

    it('should handle large counts', () => {
      store.getState().setVerticalCounts({ all: 999999 });
      expect(store.getState().verticalCounts.all).toBe(999999);
    });
  });
});
