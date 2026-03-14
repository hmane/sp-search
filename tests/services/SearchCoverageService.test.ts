import {
  SearchCoverageService,
  buildCoverageProfileFromSearchResultsConfig,
  buildCoverageQueryText,
  matchesCoverageItem,
  normalizeCoverageProfile,
  normalizeDelimitedValues
} from '../../src/libraries/spSearchStore/services/SearchCoverageService';

describe('SearchCoverageService helpers', () => {
  describe('normalizeDelimitedValues', () => {
    it('splits comma, semicolon, and newline separated values', () => {
      expect(normalizeDelimitedValues('a,b; c\nd')).toEqual(['a', 'b', 'c', 'd']);
    });

    it('deduplicates values case-insensitively', () => {
      expect(normalizeDelimitedValues('Docs,docs,DOCS')).toEqual(['Docs']);
    });
  });

  describe('normalizeCoverageProfile', () => {
    it('applies defaults and normalizes delimited scope fields', () => {
      const profile = normalizeCoverageProfile({
        title: 'Policies',
        sourceUrls: '/sites/hr/Policies,/sites/hr/Handbook',
        contentTypeIds: '0x0101;0x010100',
        excludePaths: '/sites/hr/Policies/Archive',
      });

      expect(profile.id).toBe('Policies');
      expect(profile.queryTemplate).toBe('{searchTerms}');
      expect(profile.trimDuplicates).toBe(true);
      expect(profile.includeFolders).toBe(false);
      expect(profile.sourceUrls).toEqual([
        'http://localhost/sites/hr/Policies',
        'http://localhost/sites/hr/Handbook'
      ]);
      expect(profile.contentTypeIds).toEqual(['0x0101', '0x010100']);
      expect(profile.excludePaths).toEqual(['http://localhost/sites/hr/Policies/Archive']);
    });
  });

  describe('buildCoverageQueryText', () => {
    it('builds a scoped KQL query from sources, content types, and exclusions', () => {
      const profile = normalizeCoverageProfile({
        title: 'Documents',
        sourceUrls: ['/sites/demo/Docs'],
        contentTypeIds: ['0x0101'],
        excludePaths: ['/sites/demo/Docs/Archive']
      });

      expect(buildCoverageQueryText(profile)).toBe(
        '* AND (Path:"http://localhost/sites/demo/Docs") AND (ContentTypeId:0x0101*) AND NOT Path:"http://localhost/sites/demo/Docs/Archive"'
      );
    });
  });

  describe('buildCoverageProfileFromSearchResultsConfig', () => {
    it('derives a usable profile from current-site Search Results settings', () => {
      const discovered = buildCoverageProfileFromSearchResultsConfig(
        'https://contoso.sharepoint.com/sites/search/SitePages/Search.aspx',
        {
          searchContextId: 'default',
          searchScope: 'currentsite',
          queryTemplate: '{searchTerms} IsDocument:1',
          resultSourceId: 'source-guid',
          trimDuplicates: false,
          refinementFilters: 'FileType:or("docx","pdf")'
        }
      );

      expect(discovered.profile.queryTemplate).toBe('{searchTerms} IsDocument:1');
      expect(discovered.profile.resultSourceId).toBe('source-guid');
      expect(discovered.profile.trimDuplicates).toBe(false);
      expect(discovered.profile.refinementFilters).toBe('FileType:or("docx","pdf")');
      expect(discovered.profile.sourceUrls).toEqual(['https://contoso.sharepoint.com/sites/search']);
      expect(discovered.warnings.length).toBeGreaterThan(0);
    });

    it('keeps source scope empty for all-sharepoint queries', () => {
      const discovered = buildCoverageProfileFromSearchResultsConfig(
        'https://contoso.sharepoint.com/sites/search/SitePages/Search.aspx',
        {
          searchScope: 'all',
          queryTemplate: '{searchTerms}'
        }
      );

      expect(discovered.profile.sourceUrls).toEqual([]);
      expect(discovered.warnings[0]).toContain('All SharePoint');
    });
  });

  describe('matchesCoverageItem', () => {
    it('filters out folders when includeFolders is false', () => {
      const profile = normalizeCoverageProfile({
        title: 'Files only',
        sourceUrls: ['/sites/demo/Docs']
      });

      expect(matchesCoverageItem(profile, {
        path: '/sites/demo/Docs/Folder A',
        isFolder: true
      })).toBe(false);
    });

    it('matches only configured content types and excludes excluded paths', () => {
      const profile = normalizeCoverageProfile({
        title: 'Policies',
        sourceUrls: ['/sites/demo/Policies'],
        contentTypeIds: ['0x0101'],
        excludePaths: ['/sites/demo/Policies/Archive']
      });

      expect(matchesCoverageItem(profile, {
        path: '/sites/demo/Policies/Policy.docx',
        contentTypeId: '0x010100ABC',
        isFolder: false
      })).toBe(true);

      expect(matchesCoverageItem(profile, {
        path: '/sites/demo/Policies/Archive/Old.docx',
        contentTypeId: '0x010100ABC',
        isFolder: false
      })).toBe(false);

      expect(matchesCoverageItem(profile, {
        path: '/sites/demo/Policies/Page.aspx',
        contentTypeId: '0x0120D520',
        isFolder: false
      })).toBe(false);
    });
  });

  describe('SearchCoverageService.evaluateProfile', () => {
    it('reports trimmed, untrimmed, and duplicate delta counts', async () => {
      const service: any = new SearchCoverageService();
      const profile = normalizeCoverageProfile({
        title: 'Policies',
        sourceUrls: ['/sites/demo/Policies'],
        trimDuplicates: true
      });
      const normalizedSourceUrl = 'http://localhost/sites/demo/Policies';

      service._loadSource = jest.fn().mockResolvedValue({
        title: 'Policies',
        sourceUrl: normalizedSourceUrl,
        serverRelativeUrl: '/sites/demo/Policies',
        noCrawl: false,
        hidden: false,
        sourceCount: 10,
        items: [{
          title: 'Policy.docx',
          path: normalizedSourceUrl + '/Policy.docx',
          modified: new Date('2026-01-01T00:00:00Z'),
          contentTypeId: '0x0101'
        }]
      });
      service._executeCountQuery = jest.fn().mockImplementation(function (
        _profile,
        sourceUrls: string[],
        trimDuplicates: boolean
      ): number {
        if (sourceUrls.length === 1 && sourceUrls[0] === normalizedSourceUrl) {
          return trimDuplicates ? 8 : 10;
        }
        return trimDuplicates ? 8 : 10;
      });
      service._findMissingSamples = jest.fn().mockResolvedValue([]);

      const result = await service.evaluateProfile(profile, new AbortController().signal);

      expect(result.searchCount).toBe(8);
      expect(result.searchCountTrimmed).toBe(8);
      expect(result.searchCountUntrimmed).toBe(10);
      expect(result.duplicateDelta).toBe(2);
      expect(result.delta).toBe(2);
      expect(result.sourceResults[0].searchCount).toBe(8);
      expect(result.sourceResults[0].searchCountTrimmed).toBe(8);
      expect(result.sourceResults[0].searchCountUntrimmed).toBe(10);
      expect(result.sourceResults[0].duplicateDelta).toBe(2);
      expect(result.warnings.some(function (warning): boolean {
        return warning.indexOf('Duplicate collapsing changes the indexed count by 2 items.') >= 0;
      })).toBe(true);
    });
  });
});
