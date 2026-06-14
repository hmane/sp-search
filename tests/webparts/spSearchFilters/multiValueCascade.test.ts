import { buildTagBoxBatchPayload } from '@webparts/spSearchFilters/components/TagBoxFilter';
import { buildDropdownBatchPayload } from '@webparts/spSearchFilters/components/DropdownFilter';
import { buildTaxonomyBatchPayload } from '@webparts/spSearchFilters/components/TaxonomyTreeFilter';
import type { IActiveFilter, IRefinerValue } from '@interfaces/index';

describe('TagBoxFilter batched callback', () => {
  const values: IRefinerValue[] = [
    { name: 'PDF', value: '"pdf"', count: 10, isSelected: false },
    { name: 'Word', value: '"docx"', count: 5, isSelected: false },
    { name: 'Excel', value: '"xlsx"', count: 3, isSelected: false },
  ];

  it('builds a single batched payload with all selected values, using OR operator and resolved display labels', () => {
    const payload = buildTagBoxBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"pdf"', '"docx"', '"xlsx"'],
      refinerValues: values,
      operator: 'OR',
    });

    expect(payload.filterName).toBe('FileType');
    expect(payload.values).toHaveLength(3);
    expect(payload.values[0]).toEqual<IActiveFilter>({
      filterName: 'FileType',
      value: '"pdf"',
      displayValue: 'PDF',
      operator: 'OR',
    });
    expect(payload.values[1]).toEqual<IActiveFilter>({
      filterName: 'FileType',
      value: '"docx"',
      displayValue: 'Word',
      operator: 'OR',
    });
    expect(payload.values[2]).toEqual<IActiveFilter>({
      filterName: 'FileType',
      value: '"xlsx"',
      displayValue: 'Excel',
      operator: 'OR',
    });
  });

  it('preserves selection order from the editor (not refiner bucket order)', () => {
    const payload = buildTagBoxBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"xlsx"', '"pdf"'],
      refinerValues: values,
      operator: 'OR',
    });

    expect(payload.values.map(function (v: IActiveFilter): string { return v.value; }))
      .toEqual(['"xlsx"', '"pdf"']);
  });

  it('returns an empty values array when nothing is selected (clears the filter)', () => {
    const payload = buildTagBoxBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: [],
      refinerValues: values,
      operator: 'OR',
    });

    expect(payload).toEqual({ filterName: 'FileType', values: [] });
  });

  it('leaves displayValue undefined for tokens that have no matching refiner bucket', () => {
    const payload = buildTagBoxBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"pptx"'],
      refinerValues: values,
      operator: 'AND',
    });

    expect(payload.values).toEqual<IActiveFilter[]>([
      {
        filterName: 'FileType',
        value: '"pptx"',
        displayValue: undefined,
        operator: 'AND',
      },
    ]);
  });

  it('propagates the AND operator when configured', () => {
    const payload = buildTagBoxBatchPayload({
      filterName: 'Tags',
      nextSelectedTokens: ['"a"', '"b"'],
      refinerValues: [
        { name: 'A', value: '"a"', count: 1, isSelected: false },
        { name: 'B', value: '"b"', count: 1, isSelected: false },
      ],
      operator: 'AND',
    });

    expect(payload.values.every(function (v: IActiveFilter): boolean { return v.operator === 'AND'; }))
      .toBe(true);
  });

  it('output depends only on nextSelectedTokens, not on prior activeFilters state', () => {
    // The whole point of the batched callback is that the payload carries
    // the FULL intended selection — it must NOT be computed as a delta
    // against any external state. Calling the helper with the same selection
    // tokens must always produce identical output, regardless of context.
    const independenceValues: IRefinerValue[] = [
      { name: 'pdf', value: '"pdf"', count: 10, isSelected: false },
      { name: 'docx', value: '"docx"', count: 5, isSelected: false },
      { name: 'xlsx', value: '"xlsx"', count: 3, isSelected: false },
    ];

    const result1 = buildTagBoxBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"pdf"', '"docx"', '"xlsx"'],
      refinerValues: independenceValues,
      operator: 'OR',
    });
    const result2 = buildTagBoxBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"pdf"', '"docx"', '"xlsx"'],
      refinerValues: independenceValues,
      operator: 'OR',
    });

    expect(result1).toEqual(result2);
    expect(result1.values).toHaveLength(3);
    expect(result1.values.map(function (v: IActiveFilter): string { return v.value; }))
      .toEqual(['"pdf"', '"docx"', '"xlsx"']);
  });
});

describe('DropdownFilter batched callback', () => {
  const values: IRefinerValue[] = [
    { name: 'PDF', value: '"pdf"', count: 10, isSelected: false },
    { name: 'Word', value: '"docx"', count: 5, isSelected: false },
    { name: 'Excel', value: '"xlsx"', count: 3, isSelected: false },
  ];

  it('builds a single batched payload with all selected values, using OR operator and resolved display labels', () => {
    const payload = buildDropdownBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"pdf"', '"docx"', '"xlsx"'],
      refinerValues: values,
      operator: 'OR',
    });

    expect(payload.filterName).toBe('FileType');
    expect(payload.values).toHaveLength(3);
    expect(payload.values[0]).toEqual<IActiveFilter>({
      filterName: 'FileType',
      value: '"pdf"',
      displayValue: 'PDF',
      operator: 'OR',
    });
    expect(payload.values[1]).toEqual<IActiveFilter>({
      filterName: 'FileType',
      value: '"docx"',
      displayValue: 'Word',
      operator: 'OR',
    });
    expect(payload.values[2]).toEqual<IActiveFilter>({
      filterName: 'FileType',
      value: '"xlsx"',
      displayValue: 'Excel',
      operator: 'OR',
    });
  });

  it('preserves selection order from the editor (not refiner bucket order)', () => {
    const payload = buildDropdownBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"xlsx"', '"pdf"'],
      refinerValues: values,
      operator: 'OR',
    });

    expect(payload.values.map(function (v: IActiveFilter): string { return v.value; }))
      .toEqual(['"xlsx"', '"pdf"']);
  });

  it('returns an empty values array when nothing is selected (clears the filter)', () => {
    const payload = buildDropdownBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: [],
      refinerValues: values,
      operator: 'OR',
    });

    expect(payload).toEqual({ filterName: 'FileType', values: [] });
  });

  it('leaves displayValue undefined for tokens that have no matching refiner bucket', () => {
    const payload = buildDropdownBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"pptx"'],
      refinerValues: values,
      operator: 'AND',
    });

    expect(payload.values).toEqual<IActiveFilter[]>([
      {
        filterName: 'FileType',
        value: '"pptx"',
        displayValue: undefined,
        operator: 'AND',
      },
    ]);
  });

  it('propagates the AND operator when configured', () => {
    const payload = buildDropdownBatchPayload({
      filterName: 'Tags',
      nextSelectedTokens: ['"a"', '"b"'],
      refinerValues: [
        { name: 'A', value: '"a"', count: 1, isSelected: false },
        { name: 'B', value: '"b"', count: 1, isSelected: false },
      ],
      operator: 'AND',
    });

    expect(payload.values.every(function (v: IActiveFilter): boolean { return v.operator === 'AND'; }))
      .toBe(true);
  });

  it('output depends only on nextSelectedTokens, not on prior activeFilters state', () => {
    // The whole point of the batched callback is that the payload carries
    // the FULL intended selection — it must NOT be computed as a delta
    // against any external state. Calling the helper with the same selection
    // tokens must always produce identical output, regardless of context.
    const independenceValues: IRefinerValue[] = [
      { name: 'pdf', value: '"pdf"', count: 10, isSelected: false },
      { name: 'docx', value: '"docx"', count: 5, isSelected: false },
      { name: 'xlsx', value: '"xlsx"', count: 3, isSelected: false },
    ];

    const result1 = buildDropdownBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"pdf"', '"docx"', '"xlsx"'],
      refinerValues: independenceValues,
      operator: 'OR',
    });
    const result2 = buildDropdownBatchPayload({
      filterName: 'FileType',
      nextSelectedTokens: ['"pdf"', '"docx"', '"xlsx"'],
      refinerValues: independenceValues,
      operator: 'OR',
    });

    expect(result1).toEqual(result2);
    expect(result1.values).toHaveLength(3);
    expect(result1.values.map(function (v: IActiveFilter): string { return v.value; }))
      .toEqual(['"pdf"', '"docx"', '"xlsx"']);
  });
});

describe('buildTaxonomyBatchPayload', () => {
  it('maps selected node keys to a single batched payload', () => {
    const tokenMap = new Map<string, string>([
      ['aaaa-bbbb-cccc-1', 'GP0|#aaaa-bbbb-cccc-1'],
      ['aaaa-bbbb-cccc-2', 'GP0|#aaaa-bbbb-cccc-2'],
    ]);
    const labelMap = new Map<string, string>([
      ['aaaa-bbbb-cccc-1', 'Electronics'],
      ['aaaa-bbbb-cccc-2', 'Books'],
    ]);

    const payload = buildTaxonomyBatchPayload(
      'owstaxIdProductCategory',
      ['aaaa-bbbb-cccc-1', 'aaaa-bbbb-cccc-2'],
      tokenMap,
      labelMap,
      'OR'
    );

    expect(payload).toEqual({
      filterName: 'owstaxIdProductCategory',
      values: [
        {
          filterName: 'owstaxIdProductCategory',
          value: 'GP0|#aaaa-bbbb-cccc-1',
          displayValue: 'Electronics',
          operator: 'OR',
        },
        {
          filterName: 'owstaxIdProductCategory',
          value: 'GP0|#aaaa-bbbb-cccc-2',
          displayValue: 'Books',
          operator: 'OR',
        },
      ],
    });
  });

  it('falls back to GP0|#<guid> when tokenMap has no entry for a key', () => {
    const payload = buildTaxonomyBatchPayload(
      'Cat',
      ['unmapped-guid'],
      new Map(),
      new Map(),
      'OR'
    );
    expect(payload.values[0].value).toBe('GP0|#unmapped-guid');
    expect(payload.values[0].displayValue).toBeUndefined();
  });
});
