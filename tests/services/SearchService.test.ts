import { SearchService } from '../../src/libraries/spSearchStore/services/SearchService';
import { IActiveFilter, ISortField } from '../../src/libraries/spSearchStore/interfaces';
import { ITokenContext } from '../../src/libraries/spSearchStore/services/TokenService';
import { createMockTokenContext } from '../utils/testHelpers';

describe('SearchService', () => {
  let tokenContext: ITokenContext;

  beforeEach(() => {
    tokenContext = createMockTokenContext();
  });

  describe('buildKqlQuery', () => {
    it('should resolve {searchTerms} token with the provided queryText', () => {
      const result = SearchService.buildKqlQuery('{searchTerms}', 'budget report', [], tokenContext);
      expect(result).toBe('budget report');
    });

    it('should override tokenContext.queryText with the provided queryText parameter', () => {
      // The tokenContext has "annual report" but we pass "budget" as queryText
      const result = SearchService.buildKqlQuery('{searchTerms}', 'budget', [], tokenContext);
      expect(result).toBe('budget');
    });

    it('should resolve a complex template with multiple tokens', () => {
      const template = '{searchTerms} Path:{Site.URL}';
      const result = SearchService.buildKqlQuery(template, 'docs', [], tokenContext);
      expect(result).toBe('docs Path:https://contoso.sharepoint.com/sites/intranet');
    });

    it('should trim leading and trailing whitespace', () => {
      const result = SearchService.buildKqlQuery('  {searchTerms}  ', 'hello', [], tokenContext);
      expect(result).toBe('hello');
    });

    it('should handle empty queryText', () => {
      const result = SearchService.buildKqlQuery('{searchTerms}', '', [], tokenContext);
      expect(result).toBe('');
    });

    it('should handle template with no tokens', () => {
      const result = SearchService.buildKqlQuery('contentclass:STS_Web', '', [], tokenContext);
      expect(result).toBe('contentclass:STS_Web');
    });

    it('should handle an empty template', () => {
      const result = SearchService.buildKqlQuery('', 'test', [], tokenContext);
      expect(result).toBe('');
    });
  });

  describe('buildRefinementFilters', () => {
    it('should return empty array for empty filters', () => {
      expect(SearchService.buildRefinementFilters([])).toEqual([]);
    });

    it('should return empty array for null/undefined filters', () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect(SearchService.buildRefinementFilters(null as any)).toEqual([]);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect(SearchService.buildRefinementFilters(undefined as any)).toEqual([]);
    });

    it('should produce a single filter string for one value', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
      ];
      const result = SearchService.buildRefinementFilters(filters);
      expect(result).toEqual(['FileType:"docx"']);
    });

    it('should group multiple values for the same property with or()', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
        { filterName: 'FileType', value: '"pptx"', operator: 'OR' },
      ];
      const result = SearchService.buildRefinementFilters(filters);
      expect(result).toEqual(['FileType:or("docx","pptx")']);
    });

    it('should keep different properties as separate entries', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
        { filterName: 'Author', value: '"John"', operator: 'OR' },
      ];
      const result = SearchService.buildRefinementFilters(filters);
      expect(result).toHaveLength(2);
      expect(result).toContain('FileType:"docx"');
      expect(result).toContain('Author:"John"');
    });

    it('should combine three values with or()', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
        { filterName: 'FileType', value: '"pptx"', operator: 'OR' },
        { filterName: 'FileType', value: '"pdf"', operator: 'OR' },
      ];
      const result = SearchService.buildRefinementFilters(filters);
      expect(result).toEqual(['FileType:or("docx","pptx","pdf")']);
    });

    it('should pass FQL range() through as-is', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'Created', value: 'range(2024-01-01,2024-12-31)', operator: 'AND' },
      ];
      const result = SearchService.buildRefinementFilters(filters);
      expect(result).toEqual(['Created:range(2024-01-01,2024-12-31)']);
    });

    it('should handle mixed regular and FQL values for same property', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'Size', value: 'range(0,1048576)', operator: 'OR' },
        { filterName: 'Size', value: 'range(1048576,10485760)', operator: 'OR' },
      ];
      const result = SearchService.buildRefinementFilters(filters);
      expect(result).toEqual(['Size:or(range(0,1048576),range(1048576,10485760))']);
    });

    it('should handle complex scenario with multiple properties and values', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
        { filterName: 'FileType', value: '"pptx"', operator: 'OR' },
        { filterName: 'Author', value: '"John"', operator: 'OR' },
        { filterName: 'Created', value: 'range(2024-01-01,2024-12-31)', operator: 'AND' },
      ];
      const result = SearchService.buildRefinementFilters(filters);
      expect(result).toHaveLength(3);
      expect(result[0]).toBe('FileType:or("docx","pptx")');
      expect(result[1]).toBe('Author:"John"');
      expect(result[2]).toBe('Created:range(2024-01-01,2024-12-31)');
    });
  });

  describe('buildSortList', () => {
    it('should return empty array when sort is undefined', () => {
      expect(SearchService.buildSortList(undefined)).toEqual([]);
    });

    it('should map "Descending" to Direction 1', () => {
      const sort: ISortField = { property: 'LastModifiedTime', direction: 'Descending' };
      const result = SearchService.buildSortList(sort);
      expect(result).toEqual([
        { Property: 'LastModifiedTime', Direction: 1 },
      ]);
    });

    it('should map "Ascending" to Direction 0', () => {
      const sort: ISortField = { property: 'Title', direction: 'Ascending' };
      const result = SearchService.buildSortList(sort);
      expect(result).toEqual([
        { Property: 'Title', Direction: 0 },
      ]);
    });

    it('should use the exact property name', () => {
      const sort: ISortField = { property: 'RefinableDate00', direction: 'Descending' };
      const result = SearchService.buildSortList(sort);
      expect(result[0].Property).toBe('RefinableDate00');
    });
  });

  describe('buildSelectedProperties', () => {
    it('should return default properties when no custom properties are provided', () => {
      const result = SearchService.buildSelectedProperties();
      expect(result).toContain('Title');
      expect(result).toContain('Path');
      expect(result).toContain('Author');
      expect(result).toContain('LastModifiedTime');
      expect(result).toContain('FileType');
      expect(result).toContain('HitHighlightedSummary');
      expect(result).toContain('Size');
    });

    it('should return default properties for empty array', () => {
      const result = SearchService.buildSelectedProperties([]);
      expect(result).toContain('Title');
      expect(result).toContain('Path');
    });

    it('should add custom properties to defaults', () => {
      const result = SearchService.buildSelectedProperties(['CustomProp1', 'CustomProp2']);
      expect(result).toContain('Title');
      expect(result).toContain('CustomProp1');
      expect(result).toContain('CustomProp2');
    });

    it('should deduplicate when custom property already exists in defaults', () => {
      const result = SearchService.buildSelectedProperties(['Title', 'Path', 'CustomProp']);
      // Title and Path are already in defaults — they should not appear twice
      const titleCount = result.filter(p => p === 'Title').length;
      expect(titleCount).toBe(1);
      expect(result).toContain('CustomProp');
    });

    it('should preserve order — defaults first, then custom', () => {
      const result = SearchService.buildSelectedProperties(['ZZCustom']);
      const titleIdx = result.indexOf('Title');
      const customIdx = result.indexOf('ZZCustom');
      expect(titleIdx).toBeLessThan(customIdx);
    });

    it('should return a new array (not a reference to the internal default)', () => {
      const result1 = SearchService.buildSelectedProperties();
      const result2 = SearchService.buildSelectedProperties();
      expect(result1).not.toBe(result2);
      expect(result1).toEqual(result2);
    });
  });
});
