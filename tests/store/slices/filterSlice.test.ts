import { StoreApi } from 'zustand/vanilla';
import { ISearchStore, IActiveFilter, IRefiner } from '../../../src/libraries/spSearchStore/interfaces';
import { createMockStore, createMockRefiner } from '../../utils/testHelpers';

describe('filterSlice', () => {
  let store: StoreApi<ISearchStore>;

  beforeEach(() => {
    store = createMockStore();
  });

  describe('initial state', () => {
    it('should have empty activeFilters', () => {
      expect(store.getState().activeFilters).toEqual([]);
    });

    it('should have empty availableRefiners', () => {
      expect(store.getState().availableRefiners).toEqual([]);
    });

    it('should have empty displayRefiners', () => {
      expect(store.getState().displayRefiners).toEqual([]);
    });

    it('should have empty filterConfig', () => {
      expect(store.getState().filterConfig).toEqual([]);
    });

    it('should not be refining initially', () => {
      expect(store.getState().isRefining).toBe(false);
    });
  });

  describe('setRefiner', () => {
    it('should add a new filter when none exist', () => {
      const filter: IActiveFilter = {
        filterName: 'FileType',
        value: '"docx"',
        operator: 'OR',
      };
      store.getState().setRefiner(filter);

      expect(store.getState().activeFilters).toHaveLength(1);
      expect(store.getState().activeFilters[0]).toEqual(filter);
    });

    it('should add a filter for a different property', () => {
      const filter1: IActiveFilter = {
        filterName: 'FileType',
        value: '"docx"',
        operator: 'OR',
      };
      const filter2: IActiveFilter = {
        filterName: 'Author',
        value: '"John Doe"',
        operator: 'OR',
      };

      store.getState().setRefiner(filter1);
      store.getState().setRefiner(filter2);

      expect(store.getState().activeFilters).toHaveLength(2);
    });

    it('should add a second value for the same property', () => {
      const filter1: IActiveFilter = {
        filterName: 'FileType',
        value: '"docx"',
        operator: 'OR',
      };
      const filter2: IActiveFilter = {
        filterName: 'FileType',
        value: '"pptx"',
        operator: 'OR',
      };

      store.getState().setRefiner(filter1);
      store.getState().setRefiner(filter2);

      expect(store.getState().activeFilters).toHaveLength(2);
      expect(store.getState().activeFilters[0].value).toBe('"docx"');
      expect(store.getState().activeFilters[1].value).toBe('"pptx"');
    });

    it('should toggle off (remove) an existing filter with same filterName and value', () => {
      const filter: IActiveFilter = {
        filterName: 'FileType',
        value: '"docx"',
        operator: 'OR',
      };

      // Add
      store.getState().setRefiner(filter);
      expect(store.getState().activeFilters).toHaveLength(1);

      // Toggle off — same filterName + value
      store.getState().setRefiner(filter);
      expect(store.getState().activeFilters).toHaveLength(0);
    });

    it('should only toggle the matching filter, not others with same filterName', () => {
      const docx: IActiveFilter = {
        filterName: 'FileType',
        value: '"docx"',
        operator: 'OR',
      };
      const pptx: IActiveFilter = {
        filterName: 'FileType',
        value: '"pptx"',
        operator: 'OR',
      };

      store.getState().setRefiner(docx);
      store.getState().setRefiner(pptx);
      expect(store.getState().activeFilters).toHaveLength(2);

      // Toggle off just docx
      store.getState().setRefiner(docx);

      expect(store.getState().activeFilters).toHaveLength(1);
      expect(store.getState().activeFilters[0].value).toBe('"pptx"');
    });

    it('should handle toggle with different operators as different entries', () => {
      // The toggle matches on filterName + value, not operator
      const filterOR: IActiveFilter = {
        filterName: 'FileType',
        value: '"docx"',
        operator: 'OR',
      };

      store.getState().setRefiner(filterOR);
      expect(store.getState().activeFilters).toHaveLength(1);

      // Same filterName and value — toggles off regardless of operator
      const filterAND: IActiveFilter = {
        filterName: 'FileType',
        value: '"docx"',
        operator: 'AND',
      };
      store.getState().setRefiner(filterAND);
      expect(store.getState().activeFilters).toHaveLength(0);
    });
  });

  describe('removeRefiner', () => {
    it('should remove a specific filter by filterName and value', () => {
      const filter1: IActiveFilter = { filterName: 'FileType', value: '"docx"', operator: 'OR' };
      const filter2: IActiveFilter = { filterName: 'FileType', value: '"pptx"', operator: 'OR' };
      const filter3: IActiveFilter = { filterName: 'Author', value: '"John"', operator: 'OR' };

      store.getState().setRefiner(filter1);
      store.getState().setRefiner(filter2);
      store.getState().setRefiner(filter3);
      expect(store.getState().activeFilters).toHaveLength(3);

      store.getState().removeRefiner('FileType', '"docx"');
      expect(store.getState().activeFilters).toHaveLength(2);
      expect(store.getState().activeFilters.find(f => f.value === '"docx"')).toBeUndefined();
    });

    it('should remove all filters for a property when value is omitted', () => {
      const filter1: IActiveFilter = { filterName: 'FileType', value: '"docx"', operator: 'OR' };
      const filter2: IActiveFilter = { filterName: 'FileType', value: '"pptx"', operator: 'OR' };
      const filter3: IActiveFilter = { filterName: 'Author', value: '"John"', operator: 'OR' };

      store.getState().setRefiner(filter1);
      store.getState().setRefiner(filter2);
      store.getState().setRefiner(filter3);

      store.getState().removeRefiner('FileType');

      expect(store.getState().activeFilters).toHaveLength(1);
      expect(store.getState().activeFilters[0].filterName).toBe('Author');
    });

    it('should be a no-op when the filter does not exist', () => {
      const filter: IActiveFilter = { filterName: 'FileType', value: '"docx"', operator: 'OR' };
      store.getState().setRefiner(filter);

      store.getState().removeRefiner('NonExistent', '"value"');
      expect(store.getState().activeFilters).toHaveLength(1);
    });

    it('should handle removing from an empty array', () => {
      store.getState().removeRefiner('FileType', '"docx"');
      expect(store.getState().activeFilters).toEqual([]);
    });

    it('should remove only the exact value match, keeping others', () => {
      store.getState().setRefiner({ filterName: 'Author', value: '"John"', operator: 'OR' });
      store.getState().setRefiner({ filterName: 'Author', value: '"Jane"', operator: 'OR' });

      store.getState().removeRefiner('Author', '"John"');

      expect(store.getState().activeFilters).toHaveLength(1);
      expect(store.getState().activeFilters[0].value).toBe('"Jane"');
    });
  });

  describe('clearAllFilters', () => {
    it('should remove all active filters', () => {
      store.getState().setRefiner({ filterName: 'FileType', value: '"docx"', operator: 'OR' });
      store.getState().setRefiner({ filterName: 'Author', value: '"John"', operator: 'OR' });
      store.getState().setRefiner({ filterName: 'Created', value: 'range(2024-01-01,2024-12-31)', operator: 'AND' });

      store.getState().clearAllFilters();
      expect(store.getState().activeFilters).toEqual([]);
    });

    it('should be a no-op on an already empty array', () => {
      store.getState().clearAllFilters();
      expect(store.getState().activeFilters).toEqual([]);
    });

    it('should not affect availableRefiners', () => {
      const refiners = [createMockRefiner()];
      store.getState().setAvailableRefiners(refiners);
      store.getState().setRefiner({ filterName: 'FileType', value: '"docx"', operator: 'OR' });

      store.getState().clearAllFilters();

      expect(store.getState().activeFilters).toEqual([]);
      expect(store.getState().availableRefiners).toHaveLength(1);
    });
  });

  describe('setAvailableRefiners', () => {
    it('should update availableRefiners and displayRefiners', () => {
      const refiners: IRefiner[] = [
        createMockRefiner({ filterName: 'FileType' }),
        createMockRefiner({ filterName: 'Author', values: [] }),
      ];

      store.getState().setAvailableRefiners(refiners);

      expect(store.getState().availableRefiners).toEqual(refiners);
      expect(store.getState().displayRefiners).toEqual(refiners);
    });

    it('should replace previous refiners', () => {
      const refiners1: IRefiner[] = [createMockRefiner({ filterName: 'FileType' })];
      const refiners2: IRefiner[] = [createMockRefiner({ filterName: 'Author', values: [] })];

      store.getState().setAvailableRefiners(refiners1);
      store.getState().setAvailableRefiners(refiners2);

      expect(store.getState().availableRefiners).toHaveLength(1);
      expect(store.getState().availableRefiners[0].filterName).toBe('Author');
    });

    it('should handle empty refiners array', () => {
      store.getState().setAvailableRefiners([createMockRefiner()]);
      store.getState().setAvailableRefiners([]);

      expect(store.getState().availableRefiners).toEqual([]);
      expect(store.getState().displayRefiners).toEqual([]);
    });

    it('should not affect activeFilters', () => {
      store.getState().setRefiner({ filterName: 'FileType', value: '"docx"', operator: 'OR' });
      store.getState().setAvailableRefiners([createMockRefiner()]);

      expect(store.getState().activeFilters).toHaveLength(1);
    });
  });
});
