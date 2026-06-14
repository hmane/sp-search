import { normalizeSelectedPropertyItems } from '@webparts/spSearchResults/components/selectedPropertiesConfig';

describe('normalizeSelectedPropertyItems', () => {
  it('seeds starter metadata fields when SharePoint persisted an explicit empty collection', () => {
    const result = normalizeSelectedPropertyItems([], true);

    expect(result.map((item) => item.property)).toEqual([
      'Title',
      'Author',
      'LastModifiedTime',
      'FileType',
      'Size',
      'Path',
      'SiteName',
    ]);
  });

  it('seeds starter metadata fields when only Title was persisted', () => {
    const result = normalizeSelectedPropertyItems([
      { uniqueId: 'title', property: 'Title', alias: 'Name' },
    ], true);

    expect(result.map((item) => item.property)).toEqual([
      'Title',
      'Author',
      'LastModifiedTime',
      'FileType',
      'Size',
      'Path',
      'SiteName',
    ]);
  });

  it('does not seed defaults when a non-empty custom collection normalizes to metadata fields', () => {
    const result = normalizeSelectedPropertyItems([
      { uniqueId: 'x', property: 'CustomField', alias: 'Custom' },
    ], true);

    expect(result.map((item) => item.property)).toEqual(['Title', 'CustomField']);
  });
});
