import { seedToggleDefaults } from '@store/store/storeRegistry';
import { stripDefaultToggleFilters } from '@store/utils/toggleDefaults';
import type { IActiveFilter, IFilterConfig } from '@interfaces/index';

/**
 * Tests for seedToggleDefaults — the pure helper that seeds synthetic
 * activeFilter entries for `toggle` filters whose admin-configured
 * `defaultValue` should apply on first load. URL-restore always wins
 * over the default.
 *
 * Convention notes:
 *  - Toggle stores `'1'` for true and `'0'` for false (ToggleFilter.tsx
 *    reads `'1' || 'true'` and `'0' || 'false'` for back-compat, but
 *    user-toggled writes are always `'1'` / `'0'`). Seeded entries match.
 *  - `operator` is taken from `config.operator` so seeded entries are
 *    structurally identical to user-toggled ones.
 */

const baseToggleConfig: IFilterConfig = {
  id: 'flt-confidential',
  managedProperty: 'IsConfidential',
  displayName: 'Confidential only',
  filterType: 'toggle',
  operator: 'OR',
  maxValues: 10,
  defaultExpanded: true,
  showCount: false,
  sortBy: 'count',
  sortDirection: 'desc',
  multiValues: false,
  defaultValue: true,
};

describe('seedToggleDefaults', () => {
  it('seeds a synthetic active filter when no URL state present (defaultValue=true)', () => {
    const seeded = seedToggleDefaults([], [baseToggleConfig]);
    expect(seeded).toHaveLength(1);
    expect(seeded[0]).toEqual({
      filterName: 'IsConfidential',
      value: '1',
      displayValue: 'Yes',
      operator: 'OR',
    });
  });

  it('seeds with value "0" when defaultValue=false', () => {
    const offConfig: IFilterConfig = { ...baseToggleConfig, defaultValue: false };
    const seeded = seedToggleDefaults([], [offConfig]);
    expect(seeded).toHaveLength(1);
    expect(seeded[0].value).toBe('0');
    expect(seeded[0].displayValue).toBe('No');
  });

  it('does NOT override an existing URL-restored active filter', () => {
    const urlRestored: IActiveFilter[] = [
      { filterName: 'IsConfidential', value: '0', operator: 'OR' },
    ];
    const seeded = seedToggleDefaults(urlRestored, [baseToggleConfig]);
    expect(seeded).toBe(urlRestored);
    expect(seeded).toEqual(urlRestored);
  });

  it('ignores configs without defaultValue', () => {
    const noDefault: IFilterConfig = { ...baseToggleConfig, defaultValue: undefined };
    const seeded = seedToggleDefaults([], [noDefault]);
    expect(seeded).toEqual([]);
  });

  it('ignores non-toggle filter types even with defaultValue set', () => {
    const checkboxLike: IFilterConfig = {
      ...baseToggleConfig,
      filterType: 'checkbox',
      defaultValue: true,
    };
    const seeded = seedToggleDefaults([], [checkboxLike]);
    expect(seeded).toEqual([]);
  });

  it('honours config.operator when seeding (AND vs OR)', () => {
    const andConfig: IFilterConfig = { ...baseToggleConfig, operator: 'AND' };
    const seeded = seedToggleDefaults([], [andConfig]);
    expect(seeded[0].operator).toBe('AND');
  });

  it('honours trueLabel/falseLabel + invertBoolean for displayValue', () => {
    const labelled: IFilterConfig = {
      ...baseToggleConfig,
      trueLabel: 'Confidential',
      falseLabel: 'Public',
      invertBoolean: true,
      defaultValue: true,
    };
    const seeded = seedToggleDefaults([], [labelled]);
    // invertBoolean: stored '1' but user-facing label is the false label
    expect(seeded[0].value).toBe('1');
    expect(seeded[0].displayValue).toBe('Public');
  });

  it('seeds multiple toggle configs independently', () => {
    const second: IFilterConfig = {
      ...baseToggleConfig,
      id: 'flt-archived',
      managedProperty: 'IsArchived',
      displayName: 'Archived only',
      defaultValue: false,
    };
    const seeded = seedToggleDefaults([], [baseToggleConfig, second]);
    expect(seeded).toHaveLength(2);
    expect(seeded[0].filterName).toBe('IsConfidential');
    expect(seeded[0].value).toBe('1');
    expect(seeded[1].filterName).toBe('IsArchived');
    expect(seeded[1].value).toBe('0');
  });

  it('skips toggle configs whose property is already active (mixed seed)', () => {
    const second: IFilterConfig = {
      ...baseToggleConfig,
      id: 'flt-archived',
      managedProperty: 'IsArchived',
      defaultValue: false,
    };
    const urlRestored: IActiveFilter[] = [
      { filterName: 'IsConfidential', value: '0', operator: 'OR' },
    ];
    const seeded = seedToggleDefaults(urlRestored, [baseToggleConfig, second]);
    expect(seeded).toHaveLength(2);
    expect(seeded[0]).toEqual({ filterName: 'IsConfidential', value: '0', operator: 'OR' });
    expect(seeded[1].filterName).toBe('IsArchived');
    expect(seeded[1].value).toBe('0');
  });

  it('returns the original array reference when no additions made (preserves referential equality)', () => {
    const noDefault: IFilterConfig = { ...baseToggleConfig, defaultValue: undefined };
    const current: IActiveFilter[] = [];
    const seeded = seedToggleDefaults(current, [noDefault]);
    expect(seeded).toBe(current);
  });
});

describe('stripDefaultToggleFilters', () => {
  const defaultFilter: IActiveFilter = {
    filterName: 'IsConfidential', value: '1', displayValue: 'Yes', operator: 'OR',
  };
  const overrideFilter: IActiveFilter = {
    filterName: 'IsConfidential', value: '0', displayValue: 'No', operator: 'OR',
  };
  const otherFilter: IActiveFilter = {
    filterName: 'FileType', value: 'pdf', displayValue: 'pdf', operator: 'OR',
  };

  it('removes a toggle filter sitting at its configured default value', () => {
    const result = stripDefaultToggleFilters([defaultFilter, otherFilter], [baseToggleConfig]);
    expect(result).toEqual([otherFilter]);
  });

  it('keeps a toggle filter overridden away from its default', () => {
    const result = stripDefaultToggleFilters([overrideFilter, otherFilter], [baseToggleConfig]);
    expect(result).toEqual([overrideFilter, otherFilter]);
  });

  it('round-trips with seedToggleDefaults (seed then strip yields the original)', () => {
    const seeded = seedToggleDefaults([otherFilter], [baseToggleConfig]);
    expect(seeded).toHaveLength(2);
    expect(stripDefaultToggleFilters(seeded, [baseToggleConfig])).toEqual([otherFilter]);
  });

  it('keeps everything when no toggle defaults are configured', () => {
    const noDefault: IFilterConfig = { ...baseToggleConfig, defaultValue: undefined };
    const current = [defaultFilter, otherFilter];
    expect(stripDefaultToggleFilters(current, [noDefault])).toBe(current);
  });

  it('is a no-op for empty inputs', () => {
    expect(stripDefaultToggleFilters([], [baseToggleConfig])).toEqual([]);
  });
});
