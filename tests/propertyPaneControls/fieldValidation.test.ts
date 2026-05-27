/**
 * T4.D8 — property-pane error handling parity.
 *
 * Centralised validators for the high-impact free-text fields the audit
 * called out: `coverageSourcePageUrl`, `expectedSiteUrls`,
 * `newPageQueryParameter`. The shared module also re-exports the
 * `searchContextId` required-error string (already shipped in T3.D4) so
 * downstream callers have one import surface.
 */

import {
  validateCoverageSourcePageUrl,
  validateExpectedSiteUrlsField,
  validateNewPageQueryParameter,
} from '../../src/propertyPaneControls/fieldValidation';

describe('validateCoverageSourcePageUrl', () => {
  it('passes for an empty value (field is optional)', () => {
    expect(validateCoverageSourcePageUrl('')).toBe('');
  });

  it('passes for an absolute SharePoint URL', () => {
    expect(validateCoverageSourcePageUrl('https://pixelboy.sharepoint.com/sites/SPSearch/SitePages/Search.aspx')).toBe('');
  });

  it('passes for a server-relative /sites/ path', () => {
    expect(validateCoverageSourcePageUrl('/sites/SPSearch/SitePages/Search.aspx')).toBe('');
  });

  it('passes for a server-relative /teams/ path', () => {
    expect(validateCoverageSourcePageUrl('/teams/Marketing/SitePages/Search.aspx')).toBe('');
  });

  it('flags an http (not https) absolute URL', () => {
    const msg = validateCoverageSourcePageUrl('http://pixelboy.sharepoint.com/sites/X/SitePages/Y.aspx');
    expect(msg).not.toBe('');
    expect(msg).toMatch(/https/i);
  });

  it('flags a non-SharePoint absolute URL', () => {
    const msg = validateCoverageSourcePageUrl('https://example.com/Search.aspx');
    expect(msg).not.toBe('');
    expect(msg).toMatch(/sharepoint/i);
  });

  it('flags a path that is not /sites/ or /teams/', () => {
    const msg = validateCoverageSourcePageUrl('/random/path/Search.aspx');
    expect(msg).not.toBe('');
  });
});

describe('validateExpectedSiteUrlsField', () => {
  it('passes for an empty value (field is optional)', () => {
    expect(validateExpectedSiteUrlsField('')).toBe('');
  });

  it('passes for one valid URL per line', () => {
    const v = 'https://pixelboy.sharepoint.com/sites/A\nhttps://pixelboy.sharepoint.com/sites/B';
    expect(validateExpectedSiteUrlsField(v)).toBe('');
  });

  it('reports the line number of a malformed URL', () => {
    const v = 'https://pixelboy.sharepoint.com/sites/A\njunk\nhttps://pixelboy.sharepoint.com/sites/B';
    const msg = validateExpectedSiteUrlsField(v);
    expect(msg).toMatch(/line 2/i);
  });

  it('reports the first malformed line when multiple are bad', () => {
    const v = 'good https://pixelboy.sharepoint.com/sites/A\nbad\nworse';
    const msg = validateExpectedSiteUrlsField(v);
    // The "good ..." line is itself malformed (has a space), so it should flag line 1
    expect(msg).toMatch(/line 1/i);
  });

  it('flags an http (not https) URL with the line number', () => {
    const v = 'http://pixelboy.sharepoint.com/sites/X';
    const msg = validateExpectedSiteUrlsField(v);
    expect(msg).toMatch(/https/i);
    expect(msg).toMatch(/line 1/i);
  });

  it('ignores blank lines', () => {
    const v = '\nhttps://pixelboy.sharepoint.com/sites/A\n\nhttps://pixelboy.sharepoint.com/sites/B\n';
    expect(validateExpectedSiteUrlsField(v)).toBe('');
  });
});

describe('validateNewPageQueryParameter', () => {
  it('passes for an empty value (field is optional)', () => {
    expect(validateNewPageQueryParameter('')).toBe('');
  });

  it('passes for alphanumeric', () => {
    expect(validateNewPageQueryParameter('q')).toBe('');
    expect(validateNewPageQueryParameter('searchQuery')).toBe('');
    expect(validateNewPageQueryParameter('q123')).toBe('');
  });

  it('passes for alphanumeric + dash + underscore', () => {
    expect(validateNewPageQueryParameter('search-q')).toBe('');
    expect(validateNewPageQueryParameter('search_q')).toBe('');
  });

  it('flags spaces', () => {
    const msg = validateNewPageQueryParameter('search q');
    expect(msg).not.toBe('');
  });

  it('flags special URL characters (#, &, =, ?)', () => {
    expect(validateNewPageQueryParameter('q?')).not.toBe('');
    expect(validateNewPageQueryParameter('q&')).not.toBe('');
    expect(validateNewPageQueryParameter('q=')).not.toBe('');
    expect(validateNewPageQueryParameter('q#')).not.toBe('');
  });

  it('flags a leading digit', () => {
    // Common convention — URL parameter names typically start with a letter.
    expect(validateNewPageQueryParameter('1q')).not.toBe('');
  });
});
