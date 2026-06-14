import { buildTagBoxBatchPayload } from '@webparts/spSearchFilters/components/TagBoxFilter';
import { buildDropdownBatchPayload } from '@webparts/spSearchFilters/components/DropdownFilter';
import { buildPeoplePickerBatchPayload } from '@webparts/spSearchFilters/components/PeoplePickerFilter';
import {
  buildTaxonomyTagBoxBatchPayload,
  extractGuid,
} from '@webparts/spSearchFilters/components/TaxonomyTreeFilter';
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

describe('extractGuid', () => {
  it('extracts the GUID from a well-formed GP0|#GUID token', () => {
    expect(extractGuid('GP0|#11111111-2222-3333-4444-555555555555'))
      .toBe('11111111-2222-3333-4444-555555555555');
  });

  it('is case-insensitive on hex digits', () => {
    expect(extractGuid('GP0|#AABBCCDD-EEFF-0011-2233-445566778899'))
      .toBe('AABBCCDD-EEFF-0011-2233-445566778899');
  });

  it('returns undefined for tokens without the GP0|# prefix', () => {
    expect(extractGuid('"some-value"')).toBeUndefined();
    expect(extractGuid('11111111-2222-3333-4444-555555555555')).toBeUndefined();
  });

  it('returns undefined for an empty string', () => {
    expect(extractGuid('')).toBeUndefined();
  });
});

describe('buildTaxonomyTagBoxBatchPayload', () => {
  const GUID_A = '11111111-1111-1111-1111-111111111111';
  const GUID_B = '22222222-2222-2222-2222-222222222222';
  const TOKEN_A = 'GP0|#' + GUID_A;
  const TOKEN_B = 'GP0|#' + GUID_B;

  it('maps selected taxonomy tokens to a batched payload with labels resolved from the label map', () => {
    const labelMap = new Map<string, string>([
      [GUID_A, 'Electronics'],
      [GUID_B, 'Furniture'],
    ]);

    const payload = buildTaxonomyTagBoxBatchPayload({
      filterName: 'owstaxIdCategory',
      selectedTokens: [TOKEN_A, TOKEN_B],
      labelMap,
      operator: 'OR',
    });

    expect(payload).toEqual({
      filterName: 'owstaxIdCategory',
      values: [
        {
          filterName: 'owstaxIdCategory',
          value: TOKEN_A,
          displayValue: 'Electronics',
          operator: 'OR',
        },
        {
          filterName: 'owstaxIdCategory',
          value: TOKEN_B,
          displayValue: 'Furniture',
          operator: 'OR',
        },
      ],
    });
  });

  it('falls back to the raw token as displayValue when the label is not in the map', () => {
    const payload = buildTaxonomyTagBoxBatchPayload({
      filterName: 'owstaxIdCategory',
      selectedTokens: [TOKEN_A],
      labelMap: new Map<string, string>(),
      operator: 'OR',
    });

    expect(payload.values).toEqual<IActiveFilter[]>([
      {
        filterName: 'owstaxIdCategory',
        value: TOKEN_A,
        displayValue: TOKEN_A,
        operator: 'OR',
      },
    ]);
  });

  it('looks up labels case-insensitively (labelMap is keyed by lowercase GUID)', () => {
    const labelMap = new Map<string, string>([
      [GUID_A.toLowerCase(), 'Electronics'],
    ]);
    const upperToken = 'GP0|#' + GUID_A.toUpperCase();

    const payload = buildTaxonomyTagBoxBatchPayload({
      filterName: 'owstaxIdCategory',
      selectedTokens: [upperToken],
      labelMap,
      operator: 'OR',
    });

    expect(payload.values[0].displayValue).toBe('Electronics');
  });

  it('returns an empty values array when nothing is selected (clears the filter)', () => {
    const payload = buildTaxonomyTagBoxBatchPayload({
      filterName: 'owstaxIdCategory',
      selectedTokens: [],
      labelMap: new Map<string, string>([[GUID_A, 'Electronics']]),
      operator: 'OR',
    });

    expect(payload).toEqual({ filterName: 'owstaxIdCategory', values: [] });
  });

  it('propagates the AND operator when configured', () => {
    const labelMap = new Map<string, string>([
      [GUID_A, 'Electronics'],
      [GUID_B, 'Furniture'],
    ]);

    const payload = buildTaxonomyTagBoxBatchPayload({
      filterName: 'owstaxIdCategory',
      selectedTokens: [TOKEN_A, TOKEN_B],
      labelMap,
      operator: 'AND',
    });

    expect(payload.values.every(function (v: IActiveFilter): boolean { return v.operator === 'AND'; }))
      .toBe(true);
  });

  it('preserves selection order from the editor (not labelMap iteration order)', () => {
    const labelMap = new Map<string, string>([
      [GUID_A, 'Electronics'],
      [GUID_B, 'Furniture'],
    ]);

    const payload = buildTaxonomyTagBoxBatchPayload({
      filterName: 'owstaxIdCategory',
      selectedTokens: [TOKEN_B, TOKEN_A],
      labelMap,
      operator: 'OR',
    });

    expect(payload.values.map(function (v: IActiveFilter): string { return v.value; }))
      .toEqual([TOKEN_B, TOKEN_A]);
  });

  it('output depends only on selectedTokens + labelMap, not on prior activeFilters state', () => {
    // Lock-in for the stale-state bug — same inputs must always produce
    // identical output, independent of any external/closure state.
    const labelMap = new Map<string, string>([
      [GUID_A, 'Electronics'],
      [GUID_B, 'Furniture'],
    ]);

    const result1 = buildTaxonomyTagBoxBatchPayload({
      filterName: 'owstaxIdCategory',
      selectedTokens: [TOKEN_A, TOKEN_B],
      labelMap,
      operator: 'OR',
    });
    const result2 = buildTaxonomyTagBoxBatchPayload({
      filterName: 'owstaxIdCategory',
      selectedTokens: [TOKEN_A, TOKEN_B],
      labelMap,
      operator: 'OR',
    });

    expect(result1).toEqual(result2);
    expect(result1.values).toHaveLength(2);
    expect(result1.values.map(function (v: IActiveFilter): string { return v.value; }))
      .toEqual([TOKEN_A, TOKEN_B]);
  });
});
