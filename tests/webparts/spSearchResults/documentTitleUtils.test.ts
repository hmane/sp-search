import { formatUrlBreadcrumb } from '../../../src/webparts/spSearchResults/components/documentTitleUtils';

describe('formatUrlBreadcrumb', () => {
  it('keeps the host and site path when no base URL is supplied', () => {
    expect(formatUrlBreadcrumb('https://pixelboy.sharepoint.com/sites/SPSearch/SitePages/Home.aspx'))
      .toBe('pixelboy.sharepoint.com › sites › SPSearch › SitePages');
  });

  it('omits the current site path when a matching base URL is supplied', () => {
    expect(formatUrlBreadcrumb(
      'https://pixelboy.sharepoint.com/sites/SPSearch/SitePages/Home.aspx',
      { baseUrl: 'https://pixelboy.sharepoint.com/sites/SPSearch' }
    )).toBe('SitePages');
  });

  it('shows library and folders relative to the current site', () => {
    expect(formatUrlBreadcrumb(
      'https://pixelboy.sharepoint.com/sites/SPSearch/Shared%20Documents/Client/Audit.pdf',
      { baseUrl: 'https://pixelboy.sharepoint.com/sites/SPSearch/' }
    )).toBe('Shared Documents › Client');
  });

  it('falls back to the full breadcrumb when the base URL does not match', () => {
    expect(formatUrlBreadcrumb(
      'https://pixelboy.sharepoint.com/sites/Other/SitePages/Home.aspx',
      { baseUrl: 'https://pixelboy.sharepoint.com/sites/SPSearch' }
    )).toBe('pixelboy.sharepoint.com › sites › Other › SitePages');
  });
});
