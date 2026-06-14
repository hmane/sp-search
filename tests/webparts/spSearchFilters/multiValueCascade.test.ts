import { buildTagBoxBatchPayload } from '@webparts/spSearchFilters/components/TagBoxFilter';
import { buildDropdownBatchPayload } from '@webparts/spSearchFilters/components/DropdownFilter';
import { buildPeoplePickerBatchPayload } from '@webparts/spSearchFilters/components/PeoplePickerFilter';
import type { IPersonaProps } from '@fluentui/react/lib/Persona';
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

describe('buildPeoplePickerBatchPayload', () => {
  it('maps Fluent personas to a single batched payload (value=secondaryText/claim, displayValue=text)', () => {
    const personas: IPersonaProps[] = [
      { text: 'Alice Smith', secondaryText: 'i:0#.f|membership|alice@contoso.com' },
      { text: 'Bob Jones', secondaryText: 'i:0#.f|membership|bob@contoso.com' },
    ];

    const payload = buildPeoplePickerBatchPayload({
      filterName: 'EditorOWSUSER',
      personas,
      operator: 'OR',
    });

    expect(payload).toEqual({
      filterName: 'EditorOWSUSER',
      values: [
        {
          filterName: 'EditorOWSUSER',
          value: 'i:0#.f|membership|alice@contoso.com',
          displayValue: 'Alice Smith',
          operator: 'OR',
        },
        {
          filterName: 'EditorOWSUSER',
          value: 'i:0#.f|membership|bob@contoso.com',
          displayValue: 'Bob Jones',
          operator: 'OR',
        },
      ],
    });
  });

  it('returns an empty values array when no personas are selected (clears the filter)', () => {
    const payload = buildPeoplePickerBatchPayload({
      filterName: 'EditorOWSUSER',
      personas: [],
      operator: 'OR',
    });

    expect(payload).toEqual({ filterName: 'EditorOWSUSER', values: [] });
  });
});
