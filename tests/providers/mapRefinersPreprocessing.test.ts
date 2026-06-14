import { mapRefinersWithPreprocessing } from '@providers/SharePointSearchProvider';
import type { IFilterConfig } from '@interfaces/index';

/**
 * Tests for `mapRefinersWithPreprocessing` — the pure helper extracted from
 * SharePointSearchProvider._mapRefiners that strips "type;#" prefixes (when
 * the filter's dataType warrants it) and splits delimited values into
 * separate buckets with aggregated counts.
 *
 * Key invariants:
 *   - `name` is the cleaned/split display label.
 *   - `value` preserves the original raw refinement token so the eventual
 *     KQL/FQL refinement filter still matches SharePoint's multi-value
 *     serialization (`string;#Actual Value`).
 *   - Empty cleaned strings render as "(blank)" without losing the raw token.
 */

interface IRawRefinerEntry {
  RefinementName: string;
  RefinementValue: string;
  RefinementCount: number;
}

interface IRawRefinerResponse {
  Name: string;
  Entries: IRawRefinerEntry[];
}

function refinerResponse(
  name: string,
  entries: Array<{ value: string; count: number }>
): IRawRefinerResponse {
  return {
    Name: name,
    Entries: entries.map(function (e) {
      return {
        RefinementName: e.value,
        RefinementValue: e.value,
        RefinementCount: e.count,
      };
    }),
  };
}

function cfg(overrides: Partial<IFilterConfig>): IFilterConfig {
  return {
    id: overrides.id || 'flt-' + (overrides.managedProperty || 'x'),
    managedProperty: overrides.managedProperty || 'X',
    displayName: overrides.displayName || 'X',
    filterType: overrides.filterType || 'checkbox',
    operator: overrides.operator || 'OR',
    maxValues: overrides.maxValues !== undefined ? overrides.maxValues : 10,
    defaultExpanded: overrides.defaultExpanded !== undefined ? overrides.defaultExpanded : true,
    showCount: overrides.showCount !== undefined ? overrides.showCount : true,
    sortBy: overrides.sortBy || 'count',
    sortDirection: overrides.sortDirection || 'desc',
    multiValues: overrides.multiValues !== undefined ? overrides.multiValues : true,
    ...overrides,
  };
}

describe('mapRefinersWithPreprocessing', () => {
  it('passes values through unchanged when no config exists', () => {
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('FileType', [{ value: 'pdf', count: 5 }])],
      []
    );
    expect(result[0].values[0].name).toBe('pdf');
    expect(result[0].values[0].count).toBe(5);
  });

  it('strips "string;#" prefix when dataType=choiceMulti', () => {
    const config: IFilterConfig[] = [cfg({
      managedProperty: 'DocType',
      displayName: 'Doc type',
      filterType: 'checkbox',
      dataType: 'choiceMulti',
    })];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('DocType', [
        { value: 'string;#Articles', count: 10 },
        { value: 'string;#Letters', count: 4 },
      ])],
      config
    );
    expect(result[0].values).toEqual([
      expect.objectContaining({ name: 'Articles', count: 10 }),
      expect.objectContaining({ name: 'Letters', count: 4 }),
    ]);
  });

  it('renders empty-after-strip as "(blank)" while preserving original token for KQL', () => {
    const config: IFilterConfig[] = [cfg({
      managedProperty: 'DocType',
      displayName: 'Doc type',
      filterType: 'checkbox',
      dataType: 'choiceMulti',
    })];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('DocType', [
        { value: 'string;#', count: 100 },
      ])],
      config
    );
    expect(result[0].values[0].name).toBe('(blank)');
    expect(result[0].values[0].value).toBe('string;#');
  });

  it('auto-detects "type;#" prefix when dataType is unspecified (auto default)', () => {
    const config: IFilterConfig[] = [cfg({
      managedProperty: 'Amount',
      displayName: 'Amount',
      filterType: 'checkbox',
    })];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('Amount', [
        { value: 'int;#1500', count: 3 },
        { value: 'int;#2000', count: 7 },
      ])],
      config
    );
    expect(result[0].values[0].name).toBe('1500');
    expect(result[0].values[1].name).toBe('2000');
  });

  it('does NOT strip when dataType=text even if "string;#" prefix is present', () => {
    const config: IFilterConfig[] = [cfg({
      managedProperty: 'RawText',
      displayName: 'Raw',
      filterType: 'checkbox',
      dataType: 'text',
    })];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('RawText', [{ value: 'string;#literal', count: 1 }])],
      config
    );
    expect(result[0].values[0].name).toBe('string;#literal');
  });

  it('splits on valueSplitDelimiter, aggregates counts, dedupes tokens', () => {
    const config: IFilterConfig[] = [cfg({
      managedProperty: 'Tags',
      displayName: 'Tags',
      filterType: 'tagbox',
      valueSplitDelimiter: ',',
    })];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('Tags', [
        { value: 'finance, hot, urgent', count: 5 },
        { value: 'finance, archived', count: 3 },
        { value: 'urgent', count: 2 },
      ])],
      config
    );
    const byName = new Map(result[0].values.map(function (v) { return [v.name, v.count]; }));
    expect(byName.get('finance')).toBe(8);
    expect(byName.get('hot')).toBe(5);
    expect(byName.get('urgent')).toBe(7);
    expect(byName.get('archived')).toBe(3);
  });

  it('splits on newline AND strips prefix when both are configured', () => {
    const config: IFilterConfig[] = [cfg({
      managedProperty: 'MultiTags',
      displayName: 'MultiTags',
      filterType: 'tagbox',
      dataType: 'choiceMulti',
      valueSplitDelimiter: '\n',
    })];
    const result = mapRefinersWithPreprocessing(
      [refinerResponse('MultiTags', [
        { value: 'string;#alpha\nstring;#beta', count: 4 },
      ])],
      config
    );
    const byName = new Map(result[0].values.map(function (v) { return [v.name, v.count]; }));
    expect(byName.get('alpha')).toBe(4);
    expect(byName.get('beta')).toBe(4);
  });
});
