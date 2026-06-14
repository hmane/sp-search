import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '../../src/libraries/spSearchStore/interfaces';
import { createMockStore } from '../utils/testHelpers';

describe('filter refiner cascade display', () => {
  let store: StoreApi<ISearchStore>;

  beforeEach(() => {
    store = createMockStore();
  });

  it('removes unselected values omitted by a cascaded refiner response', () => {
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
        ],
      },
    ]);
  });

  it('keeps selected values as zero-count options when SharePoint omits them', () => {
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
      filterName: 'DocumentType',
      value: '"Invoice"',
      displayValue: 'Invoice',
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
