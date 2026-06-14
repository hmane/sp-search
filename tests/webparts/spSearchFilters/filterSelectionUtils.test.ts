import { getSelectedRefinerTokens } from '@webparts/spSearchFilters/components/filterSelectionUtils';
import type { IActiveFilter, IRefinerValue } from '@interfaces/index';

describe('getSelectedRefinerTokens', () => {
  it('keeps an active selection when cascade response omits its bucket', () => {
    const values: IRefinerValue[] = [
      { name: 'Invoice', value: '"Invoice"', count: 8, isSelected: false },
    ];
    const activeFilters: IActiveFilter[] = [
      { filterName: 'DocumentType', value: '"Contract"', displayValue: 'Contract', operator: 'OR' },
    ];

    expect(getSelectedRefinerTokens('DocumentType', values, activeFilters)).toEqual(['"Contract"']);
  });

  it('uses the current bucket token when the active selection matches by display value', () => {
    const values: IRefinerValue[] = [
      { name: 'Contract', value: '"Contract"', count: 0, isSelected: false },
    ];
    const activeFilters: IActiveFilter[] = [
      { filterName: 'DocumentType', value: 'contract', displayValue: 'Contract', operator: 'OR' },
    ];

    expect(getSelectedRefinerTokens('DocumentType', values, activeFilters)).toEqual(['"Contract"']);
  });
});
