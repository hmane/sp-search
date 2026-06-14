import {
  applyReplaceRefinerValues,
  shouldCommitActiveFilters,
} from '@webparts/spSearchFilters/components/SpSearchFilters';
import type { IActiveFilter } from '@interfaces/index';

describe('applyReplaceRefinerValues', () => {
  const base: IActiveFilter[] = [
    { filterName: 'FileType', value: 'pdf', operator: 'OR' },
    { filterName: 'Author', value: 'jdoe', operator: 'OR' },
  ];

  it('replaces all values for the named filter and preserves the rest', () => {
    const next = applyReplaceRefinerValues(base, 'FileType', [
      { filterName: 'FileType', value: 'docx', operator: 'OR' },
      { filterName: 'FileType', value: 'xlsx', operator: 'OR' },
    ]);
    expect(next).toEqual([
      { filterName: 'Author', value: 'jdoe', operator: 'OR' },
      { filterName: 'FileType', value: 'docx', operator: 'OR' },
      { filterName: 'FileType', value: 'xlsx', operator: 'OR' },
    ]);
  });

  it('returns a new array reference even when contents are equivalent', () => {
    const next = applyReplaceRefinerValues(base, 'FileType', [
      { filterName: 'FileType', value: 'pdf', operator: 'OR' },
    ]);
    expect(next).not.toBe(base);
  });

  it('clears the filter when values is empty', () => {
    const next = applyReplaceRefinerValues(base, 'FileType', []);
    expect(next).toEqual([{ filterName: 'Author', value: 'jdoe', operator: 'OR' }]);
  });

  it('ignores values whose filterName does not match the target', () => {
    const next = applyReplaceRefinerValues(base, 'FileType', [
      { filterName: 'FileType', value: 'docx', operator: 'OR' },
      { filterName: 'BogusName', value: 'should-be-ignored', operator: 'OR' },
    ]);
    expect(next).toEqual([
      { filterName: 'Author', value: 'jdoe', operator: 'OR' },
      { filterName: 'FileType', value: 'docx', operator: 'OR' },
    ]);
  });
});

describe('shouldCommitActiveFilters', () => {
  it('blocks a no-op replacement from committing a fresh activeFilters array', () => {
    const current: IActiveFilter[] = [
      { filterName: 'DocumentType', value: '"Contract"', displayValue: 'Contract', operator: 'OR' },
      { filterName: 'Inactive', value: 'true', displayValue: 'Inactive', operator: 'OR' },
    ];
    const next: IActiveFilter[] = [
      { filterName: 'Inactive', value: 'true', displayValue: 'Inactive', operator: 'OR' },
      { filterName: 'DocumentType', value: '"Contract"', displayValue: 'Contract', operator: 'OR' },
    ];

    expect(next).not.toBe(current);
    expect(shouldCommitActiveFilters(current, next)).toBe(false);
  });

  it('allows real selection changes through', () => {
    const current: IActiveFilter[] = [
      { filterName: 'DocumentType', value: '"Contract"', displayValue: 'Contract', operator: 'OR' },
    ];
    const next: IActiveFilter[] = [
      { filterName: 'DocumentType', value: '"Contract"', displayValue: 'Contract', operator: 'OR' },
      { filterName: 'DocumentType', value: '"Invoice"', displayValue: 'Invoice', operator: 'OR' },
    ];

    expect(shouldCommitActiveFilters(current, next)).toBe(true);
  });
});
