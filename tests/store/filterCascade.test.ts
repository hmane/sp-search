import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '../../src/libraries/spSearchStore/interfaces';
import { createMockStore } from '../utils/testHelpers';

describe('filter refiner cascade display', () => {
  let store: StoreApi<ISearchStore>;

  beforeEach(() => {
    store = createMockStore();
  });

  it('zeros missing refiner values while active filters are applied', () => {
    store.getState().setAvailableRefiners([
      {
        filterName: 'DocumentType',
        values: [
          { name: 'Policy', value: '"Policy"', count: 12, isSelected: false },
          { name: 'Invoice', value: '"Invoice"', count: 8, isSelected: false },
        ],
      },
    ]);

    store.getState().setRefiner({
      filterName: 'Inactive',
      value: '"1"',
      displayValue: 'Yes',
      operator: 'OR',
    });

    store.getState().setAvailableRefiners([
      {
        filterName: 'DocumentType',
        values: [
          { name: 'Policy', value: '"Policy"', count: 3, isSelected: false },
        ],
      },
    ]);

    expect(store.getState().displayRefiners).toEqual([
      {
        filterName: 'DocumentType',
        values: [
          { name: 'Policy', value: '"Policy"', count: 3, isSelected: false },
          { name: 'Invoice', value: '"Invoice"', count: 0, isSelected: false },
        ],
      },
    ]);
  });
});
