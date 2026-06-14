import type { ISearchResult } from '../../../src/libraries/spSearchStore/interfaces/ISearchResult';
import { buildRowActionMenu } from '../../../src/webparts/spSearchResults/components/buildRowActionMenu';

function makeItem(): ISearchResult {
  return {
    key: '1',
    title: 'Document',
    url: 'https://contoso.sharepoint.com/sites/docs/a.pdf',
    summary: '',
    author: { displayText: '', email: '' },
    created: '',
    modified: '',
    fileType: 'pdf',
    fileSize: 0,
    siteName: '',
    siteUrl: '',
    thumbnailUrl: '',
    properties: {},
  };
}

describe('buildRowActionMenu', () => {
  it('adds collection action between open and utility actions when supplied', () => {
    const items = buildRowActionMenu(makeItem(), {
      onAddToCollection: (): void => { /* noop */ },
    });

    expect(items.map((item) => item.key)).toEqual([
      'open',
      'addToCollection',
      'download',
      'copyLink',
    ]);
  });

  it('omits collection action when no dialog trigger is supplied', () => {
    const items = buildRowActionMenu(makeItem());

    expect(items.map((item) => item.key)).toEqual([
      'open',
      'download',
      'copyLink',
    ]);
  });
});
