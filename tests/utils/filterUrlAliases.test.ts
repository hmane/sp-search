import {
  assignFilterUrlAliases,
  getFilterUrlAlias,
} from '../../src/libraries/spSearchStore/utils/filterUrlAliases';
import type { IFilterConfig } from '../../src/libraries/spSearchStore/interfaces/IFilterTypes';

/**
 * T3.D3 — `Author` + `AuthorOWSUSER` configured together produce
 * disambiguated URL aliases. Both managed properties resolve to `au`
 * via `getDefaultFilterUrlAlias`; without disambiguation, the second
 * filter's `?au=` value clobbers the first on deep-link round-trip.
 */

function cfg(overrides: Partial<IFilterConfig>): IFilterConfig {
  return {
    id: overrides.id || overrides.managedProperty || 'id',
    displayName: overrides.displayName || 'Display',
    managedProperty: overrides.managedProperty || 'FileType',
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

describe('assignFilterUrlAliases — T3.D3', () => {
  it('returns an empty map for an empty config list', () => {
    expect(assignFilterUrlAliases([])).toEqual(new Map());
  });

  it('keeps the natural alias when there is no collision', () => {
    const map = assignFilterUrlAliases([
      cfg({ id: 'f1', managedProperty: 'FileType' }),
      cfg({ id: 'f2', managedProperty: 'ContentType' }),
    ]);
    expect(map.get('f1')).toBe('ft');
    expect(map.get('f2')).toBe('ct');
  });

  it('disambiguates Author + AuthorOWSUSER (the audit acceptance case)', () => {
    const map = assignFilterUrlAliases([
      cfg({ id: 'f1', managedProperty: 'Author', filterType: 'people' }),
      cfg({ id: 'f2', managedProperty: 'AuthorOWSUSER', filterType: 'people' }),
    ]);
    expect(map.get('f1')).toBe('au');
    expect(map.get('f2')).toBe('au2');
  });

  it('produces sequential suffixes for 3+ collisions', () => {
    const map = assignFilterUrlAliases([
      cfg({ id: 'f1', managedProperty: 'Author', filterType: 'people' }),
      cfg({ id: 'f2', managedProperty: 'AuthorOWSUSER', filterType: 'people' }),
      cfg({ id: 'f3', managedProperty: 'DisplayAuthor', filterType: 'people' }),
    ]);
    expect(map.get('f1')).toBe('au');
    expect(map.get('f2')).toBe('au2');
    expect(map.get('f3')).toBe('au3');
  });

  it('admin-supplied urlAlias wins over the default but still disambiguates', () => {
    const map = assignFilterUrlAliases([
      cfg({ id: 'f1', managedProperty: 'CustomA', urlAlias: 'tag' }),
      cfg({ id: 'f2', managedProperty: 'CustomB', urlAlias: 'tag' }),
    ]);
    expect(map.get('f1')).toBe('tag');
    expect(map.get('f2')).toBe('tag2');
  });

  it('admin-supplied alias does NOT collide with a downstream default that happens to match', () => {
    // First filter explicitly aliased to 'ft'; second is FileType (default 'ft').
    const map = assignFilterUrlAliases([
      cfg({ id: 'f1', managedProperty: 'CustomA', urlAlias: 'ft' }),
      cfg({ id: 'f2', managedProperty: 'FileType' }),
    ]);
    expect(map.get('f1')).toBe('ft');
    expect(map.get('f2')).toBe('ft2');
  });

  it('first-come-first-served — order of the configs list determines who keeps the base alias', () => {
    const a = assignFilterUrlAliases([
      cfg({ id: 'f1', managedProperty: 'Author', filterType: 'people' }),
      cfg({ id: 'f2', managedProperty: 'AuthorOWSUSER', filterType: 'people' }),
    ]);
    const b = assignFilterUrlAliases([
      cfg({ id: 'f2', managedProperty: 'AuthorOWSUSER', filterType: 'people' }),
      cfg({ id: 'f1', managedProperty: 'Author', filterType: 'people' }),
    ]);
    // In `a`, f1 is first → keeps 'au'. In `b`, f2 is first → keeps 'au'.
    expect(a.get('f1')).toBe('au');
    expect(a.get('f2')).toBe('au2');
    expect(b.get('f2')).toBe('au');
    expect(b.get('f1')).toBe('au2');
  });

  it('produces the same alias deterministically across calls', () => {
    const configs = [
      cfg({ id: 'f1', managedProperty: 'Author', filterType: 'people' }),
      cfg({ id: 'f2', managedProperty: 'AuthorOWSUSER', filterType: 'people' }),
    ];
    const a = assignFilterUrlAliases(configs);
    const b = assignFilterUrlAliases(configs);
    expect(Array.from(a.entries())).toEqual(Array.from(b.entries()));
  });
});

describe('getFilterUrlAlias — back-compat', () => {
  it('still returns the natural alias for a single config (no list context)', () => {
    expect(getFilterUrlAlias({ managedProperty: 'Author', filterType: 'people', urlAlias: undefined })).toBe('au');
    expect(getFilterUrlAlias({ managedProperty: 'FileType', filterType: 'checkbox', urlAlias: undefined })).toBe('ft');
    expect(getFilterUrlAlias({ managedProperty: 'CustomX', filterType: 'checkbox', urlAlias: 'tag' })).toBe('tag');
  });
});
