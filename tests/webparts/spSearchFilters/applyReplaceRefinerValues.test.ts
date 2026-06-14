import { applyReplaceRefinerValues } from '@webparts/spSearchFilters/components/SpSearchFilters';
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
