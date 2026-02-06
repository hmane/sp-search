import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '../../../src/libraries/spSearchStore/interfaces';
import { createMockStore, createMockSearchResult } from '../../utils/testHelpers';

describe('resultSlice', () => {
  let store: StoreApi<ISearchStore>;

  beforeEach(() => {
    store = createMockStore();
  });

  describe('initial state', () => {
    it('should have empty items array', () => {
      expect(store.getState().items).toEqual([]);
    });

    it('should have totalCount of 0', () => {
      expect(store.getState().totalCount).toBe(0);
    });

    it('should start on page 1', () => {
      expect(store.getState().currentPage).toBe(1);
    });

    it('should have default pageSize of 25', () => {
      expect(store.getState().pageSize).toBe(25);
    });

    it('should have undefined sort', () => {
      expect(store.getState().sort).toBeUndefined();
    });

    it('should have empty promotedResults', () => {
      expect(store.getState().promotedResults).toEqual([]);
    });

    it('should not be loading initially', () => {
      expect(store.getState().isLoading).toBe(false);
    });

    it('should have undefined error', () => {
      expect(store.getState().error).toBeUndefined();
    });
  });

  describe('setResults', () => {
    it('should update items and totalCount', () => {
      const results = [
        createMockSearchResult({ key: '1', title: 'Result 1' }),
        createMockSearchResult({ key: '2', title: 'Result 2' }),
      ];

      store.getState().setResults(results, 42);

      expect(store.getState().items).toHaveLength(2);
      expect(store.getState().items[0].title).toBe('Result 1');
      expect(store.getState().totalCount).toBe(42);
    });

    it('should clear isLoading and error when results are set', () => {
      // Simulate a loading state
      store.getState().setLoading(true);
      store.getState().setError('Some previous error');

      const results = [createMockSearchResult()];
      store.getState().setResults(results, 1);

      expect(store.getState().isLoading).toBe(false);
      expect(store.getState().error).toBeUndefined();
    });

    it('should handle empty results', () => {
      store.getState().setResults([], 0);

      expect(store.getState().items).toEqual([]);
      expect(store.getState().totalCount).toBe(0);
    });

    it('should replace previous results', () => {
      const results1 = [createMockSearchResult({ key: '1', title: 'Old Result' })];
      const results2 = [createMockSearchResult({ key: '2', title: 'New Result' })];

      store.getState().setResults(results1, 10);
      store.getState().setResults(results2, 5);

      expect(store.getState().items).toHaveLength(1);
      expect(store.getState().items[0].title).toBe('New Result');
      expect(store.getState().totalCount).toBe(5);
    });

    it('should handle results with child results (collapsed groups)', () => {
      const childResult = createMockSearchResult({ key: 'child-1', title: 'Child Doc' });
      const parentResult = createMockSearchResult({
        key: 'parent-1',
        title: 'Parent Doc',
        isCollapsedGroup: true,
        childResults: [childResult],
        groupCount: 3,
      });

      store.getState().setResults([parentResult], 10);

      expect(store.getState().items[0].isCollapsedGroup).toBe(true);
      expect(store.getState().items[0].childResults).toHaveLength(1);
      expect(store.getState().items[0].groupCount).toBe(3);
    });
  });

  describe('setPage', () => {
    it('should update currentPage', () => {
      store.getState().setPage(3);
      expect(store.getState().currentPage).toBe(3);
    });

    it('should handle page 1', () => {
      store.getState().setPage(5);
      store.getState().setPage(1);
      expect(store.getState().currentPage).toBe(1);
    });

    it('should handle large page numbers', () => {
      store.getState().setPage(999);
      expect(store.getState().currentPage).toBe(999);
    });
  });

  describe('setSort', () => {
    it('should update sort field', () => {
      store.getState().setSort({ property: 'LastModifiedTime', direction: 'Descending' });

      expect(store.getState().sort).toEqual({
        property: 'LastModifiedTime',
        direction: 'Descending',
      });
    });

    it('should reset currentPage to 1 when sort changes', () => {
      store.getState().setPage(5);
      store.getState().setSort({ property: 'Title', direction: 'Ascending' });

      expect(store.getState().currentPage).toBe(1);
    });

    it('should replace previous sort', () => {
      store.getState().setSort({ property: 'Title', direction: 'Ascending' });
      store.getState().setSort({ property: 'Created', direction: 'Descending' });

      expect(store.getState().sort).toEqual({
        property: 'Created',
        direction: 'Descending',
      });
    });
  });

  describe('setPromotedResults', () => {
    it('should update promoted results', () => {
      const promoted = [
        { title: 'Company Portal', url: 'https://portal.contoso.com', description: 'Main portal' },
        { title: 'HR Site', url: 'https://contoso.sharepoint.com/sites/hr' },
      ];

      store.getState().setPromotedResults(promoted);
      expect(store.getState().promotedResults).toHaveLength(2);
      expect(store.getState().promotedResults[0].title).toBe('Company Portal');
    });

    it('should handle empty promoted results', () => {
      store.getState().setPromotedResults([{ title: 'Test', url: 'https://test.com' }]);
      store.getState().setPromotedResults([]);
      expect(store.getState().promotedResults).toEqual([]);
    });
  });

  describe('setLoading', () => {
    it('should set isLoading to true', () => {
      store.getState().setLoading(true);
      expect(store.getState().isLoading).toBe(true);
    });

    it('should set isLoading to false', () => {
      store.getState().setLoading(true);
      store.getState().setLoading(false);
      expect(store.getState().isLoading).toBe(false);
    });
  });

  describe('setError', () => {
    it('should set error message and clear isLoading', () => {
      store.getState().setLoading(true);
      store.getState().setError('Network timeout');

      expect(store.getState().error).toBe('Network timeout');
      expect(store.getState().isLoading).toBe(false);
    });

    it('should clear error when set to undefined', () => {
      store.getState().setError('Some error');
      store.getState().setError(undefined);

      expect(store.getState().error).toBeUndefined();
      expect(store.getState().isLoading).toBe(false);
    });

    it('should handle empty string error', () => {
      store.getState().setError('');
      expect(store.getState().error).toBe('');
    });
  });
});
