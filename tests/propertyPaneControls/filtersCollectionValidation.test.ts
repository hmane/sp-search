import {
  validateRefinerUrlAliases,
  type IFiltersCollectionItem,
} from '../../src/propertyPaneControls/filtersCollection/FiltersCollectionControl';

function refiner(overrides: Partial<IFiltersCollectionItem>): IFiltersCollectionItem {
  return {
    uniqueId: overrides.uniqueId || overrides.managedProperty || 'id',
    managedProperty: overrides.managedProperty || 'FileType',
    displayName: overrides.displayName || overrides.managedProperty || 'File Type',
    urlAlias: overrides.urlAlias,
    filterType: overrides.filterType || 'checkbox',
    operator: 'OR',
    maxValues: 10,
    defaultExpanded: true,
    showCount: true,
    sortBy: 'count',
    sortDirection: 'desc',
    multiValues: true,
  };
}

describe('validateRefinerUrlAliases', () => {
  it('flags duplicate explicit aliases after sanitization', () => {
    const issues = validateRefinerUrlAliases([
      refiner({ managedProperty: 'CustomA', displayName: 'A', urlAlias: 'Tag' }),
      refiner({ managedProperty: 'CustomB', displayName: 'B', urlAlias: 'tag!' }),
    ]);

    expect(issues).toEqual([
      { alias: 'tag', refinerNames: ['A', 'B'] },
    ]);
  });

  it('flags duplicate effective aliases generated from managed property defaults', () => {
    const issues = validateRefinerUrlAliases([
      refiner({ managedProperty: 'Author', displayName: 'Author', filterType: 'people' }),
      refiner({ managedProperty: 'AuthorOWSUSER', displayName: 'Document author', filterType: 'people' }),
    ]);

    expect(issues).toEqual([
      { alias: 'au', refinerNames: ['Author', 'Document author'] },
    ]);
  });

  it('allows unique explicit aliases', () => {
    const issues = validateRefinerUrlAliases([
      refiner({ managedProperty: 'Author', displayName: 'Author', filterType: 'people', urlAlias: 'author' }),
      refiner({ managedProperty: 'AuthorOWSUSER', displayName: 'Document author', filterType: 'people', urlAlias: 'docauthor' }),
    ]);

    expect(issues).toEqual([]);
  });
});
