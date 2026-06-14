import { buildTagBoxBatchPayload } from '@webparts/spSearchFilters/components/TagBoxFilter';
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
});
