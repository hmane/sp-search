import {
  serializeToUrl,
  deserializeFromUrl,
} from '../../src/libraries/spSearchStore/store/middleware/urlSyncMiddleware';
import { IActiveFilter, ISortField } from '../../src/libraries/spSearchStore/interfaces';

/**
 * Tests for the URL sync middleware — serialization and deserialization
 * of search state to/from URL query parameters.
 *
 * These tests mock window.location to simulate browser behavior in jsdom.
 */
describe('urlSyncMiddleware', () => {
  const originalLocation = window.location;

  /**
   * Helper to set the window.location.search for deserialization tests.
   */
  function setLocationSearch(search: string): void {
    Object.defineProperty(window, 'location', {
      value: {
        ...originalLocation,
        search,
        pathname: '/sites/intranet/SitePages/Search.aspx',
        hash: '',
      },
      writable: true,
      configurable: true,
    });
  }

  afterEach(() => {
    // Restore original location
    Object.defineProperty(window, 'location', {
      value: originalLocation,
      writable: true,
      configurable: true,
    });
  });

  describe('serializeToUrl', () => {
    beforeEach(() => {
      setLocationSearch('');
    });

    it('should produce state version tag (sv=1)', () => {
      const result = serializeToUrl({});
      expect(result).toContain('sv=1');
    });

    it('should serialize queryText as q parameter', () => {
      const result = serializeToUrl({ queryText: 'annual report' });
      const params = new URLSearchParams(result);
      expect(params.get('q')).toBe('annual report');
    });

    it('should omit q when queryText is empty', () => {
      const result = serializeToUrl({ queryText: '' });
      const params = new URLSearchParams(result);
      expect(params.has('q')).toBe(false);
    });

    it('should serialize currentVerticalKey as v parameter', () => {
      const result = serializeToUrl({ currentVerticalKey: 'documents' });
      const params = new URLSearchParams(result);
      expect(params.get('v')).toBe('documents');
    });

    it('should serialize sort as property:direction format', () => {
      const sort: ISortField = { property: 'LastModifiedTime', direction: 'Descending' };
      const result = serializeToUrl({ sort });
      const params = new URLSearchParams(result);
      expect(params.get('s')).toBe('LastModifiedTime:Descending');
    });

    it('should omit sort when undefined', () => {
      const result = serializeToUrl({ sort: undefined });
      const params = new URLSearchParams(result);
      expect(params.has('s')).toBe(false);
    });

    it('should serialize currentPage as p only when > 1', () => {
      const result = serializeToUrl({ currentPage: 3 });
      const params = new URLSearchParams(result);
      expect(params.get('p')).toBe('3');
    });

    it('should omit p when page is 1', () => {
      const result = serializeToUrl({ currentPage: 1 });
      const params = new URLSearchParams(result);
      expect(params.has('p')).toBe(false);
    });

    it('should serialize scope as sc parameter', () => {
      const result = serializeToUrl({
        scope: { id: 'currentsite', label: 'Current Site' },
      });
      const params = new URLSearchParams(result);
      expect(params.get('sc')).toBe('currentsite');
    });

    it('should serialize activeLayoutKey as l only when not "list"', () => {
      const result = serializeToUrl({ activeLayoutKey: 'grid' });
      const params = new URLSearchParams(result);
      expect(params.get('l')).toBe('grid');
    });

    it('should omit l when layout is "list" (default)', () => {
      const result = serializeToUrl({ activeLayoutKey: 'list' });
      const params = new URLSearchParams(result);
      expect(params.has('l')).toBe(false);
    });

    it('should serialize activeFilters as base64-encoded JSON (f parameter)', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
      ];
      const result = serializeToUrl({ activeFilters: filters });
      const params = new URLSearchParams(result);
      const encoded = params.get('f');

      expect(encoded).toBeTruthy();

      // Decode and verify
      const decoded = decodeURIComponent(escape(atob(encoded!)));
      const parsed = JSON.parse(decoded);
      expect(parsed).toEqual(filters);
    });

    it('should omit f when activeFilters is empty', () => {
      const result = serializeToUrl({ activeFilters: [] });
      const params = new URLSearchParams(result);
      expect(params.has('f')).toBe(false);
    });

    it('should serialize a full state with all parameters', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
      ];
      const result = serializeToUrl({
        queryText: 'budget',
        activeFilters: filters,
        currentVerticalKey: 'documents',
        sort: { property: 'Created', direction: 'Ascending' },
        currentPage: 2,
        scope: { id: 'hr', label: 'HR Site' },
        activeLayoutKey: 'card',
      });

      const params = new URLSearchParams(result);
      expect(params.get('q')).toBe('budget');
      expect(params.get('v')).toBe('documents');
      expect(params.get('s')).toBe('Created:Ascending');
      expect(params.get('p')).toBe('2');
      expect(params.get('sc')).toBe('hr');
      expect(params.get('l')).toBe('card');
      expect(params.has('f')).toBe(true);
      expect(params.get('sv')).toBe('1');
    });

    describe('with prefix (multi-context)', () => {
      it('should prefix parameter names', () => {
        const result = serializeToUrl({ queryText: 'test' }, 'ctx1');
        const params = new URLSearchParams(result);
        expect(params.get('ctx1.q')).toBe('test');
        expect(params.get('ctx1.sv')).toBe('1');
      });
    });
  });

  describe('deserializeFromUrl', () => {
    it('should return empty object when no sv tag is present', () => {
      setLocationSearch('?q=test');
      const result = deserializeFromUrl();
      expect(result).toEqual({});
    });

    it('should parse queryText from q parameter', () => {
      setLocationSearch('?sv=1&q=annual+report');
      const result = deserializeFromUrl();
      expect(result.queryText).toBe('annual report');
    });

    it('should parse currentVerticalKey from v parameter', () => {
      setLocationSearch('?sv=1&v=documents');
      const result = deserializeFromUrl();
      expect(result.currentVerticalKey).toBe('documents');
    });

    it('should parse sort from s parameter (property:direction)', () => {
      setLocationSearch('?sv=1&s=LastModifiedTime:Descending');
      const result = deserializeFromUrl();
      expect(result.sort).toEqual({
        property: 'LastModifiedTime',
        direction: 'Descending',
      });
    });

    it('should parse sort with Ascending direction', () => {
      setLocationSearch('?sv=1&s=Title:Ascending');
      const result = deserializeFromUrl();
      expect(result.sort).toEqual({
        property: 'Title',
        direction: 'Ascending',
      });
    });

    it('should ignore sort with invalid direction', () => {
      setLocationSearch('?sv=1&s=Title:InvalidDir');
      const result = deserializeFromUrl();
      expect(result.sort).toBeUndefined();
    });

    it('should parse currentPage from p parameter', () => {
      setLocationSearch('?sv=1&p=5');
      const result = deserializeFromUrl();
      expect(result.currentPage).toBe(5);
    });

    it('should ignore non-numeric page values', () => {
      setLocationSearch('?sv=1&p=abc');
      const result = deserializeFromUrl();
      expect(result.currentPage).toBeUndefined();
    });

    it('should ignore page values less than 1', () => {
      setLocationSearch('?sv=1&p=0');
      const result = deserializeFromUrl();
      expect(result.currentPage).toBeUndefined();
    });

    it('should parse scope from sc parameter', () => {
      setLocationSearch('?sv=1&sc=currentsite');
      const result = deserializeFromUrl();
      expect(result.scope).toBe('currentsite');
    });

    it('should parse activeLayoutKey from l parameter', () => {
      setLocationSearch('?sv=1&l=grid');
      const result = deserializeFromUrl();
      expect(result.activeLayoutKey).toBe('grid');
    });

    it('should parse activeFilters from base64-encoded f parameter', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
        { filterName: 'Author', value: '"John"', operator: 'OR' },
      ];
      const json = JSON.stringify(filters);
      const encoded = btoa(unescape(encodeURIComponent(json)));
      setLocationSearch(`?sv=1&f=${encoded}`);

      const result = deserializeFromUrl();
      expect(result.activeFilters).toEqual(filters);
    });

    it('should ignore malformed base64 in f parameter', () => {
      setLocationSearch('?sv=1&f=!!!not-valid-base64!!!');
      const result = deserializeFromUrl();
      expect(result.activeFilters).toBeUndefined();
    });

    it('should ignore invalid JSON in f parameter', () => {
      // Valid base64 but invalid JSON
      const encoded = btoa('not json at all');
      setLocationSearch(`?sv=1&f=${encoded}`);
      const result = deserializeFromUrl();
      expect(result.activeFilters).toBeUndefined();
    });

    it('should validate filter shape — reject items missing required fields', () => {
      const invalidFilters = [
        { filterName: 'FileType' }, // missing value and operator
      ];
      const json = JSON.stringify(invalidFilters);
      const encoded = btoa(unescape(encodeURIComponent(json)));
      setLocationSearch(`?sv=1&f=${encoded}`);

      const result = deserializeFromUrl();
      // The invalid items should be filtered out, resulting in no filters
      expect(result.activeFilters).toBeUndefined();
    });

    it('should validate filter operator — only accept AND or OR', () => {
      const filtersWithBadOperator = [
        { filterName: 'FileType', value: '"docx"', operator: 'NOT' },
      ];
      const json = JSON.stringify(filtersWithBadOperator);
      const encoded = btoa(unescape(encodeURIComponent(json)));
      setLocationSearch(`?sv=1&f=${encoded}`);

      const result = deserializeFromUrl();
      expect(result.activeFilters).toBeUndefined();
    });

    it('should parse a full URL with all parameters', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"pdf"', operator: 'OR' },
      ];
      const json = JSON.stringify(filters);
      const encoded = btoa(unescape(encodeURIComponent(json)));
      setLocationSearch(
        `?sv=1&q=report&v=documents&s=Created:Descending&p=3&sc=hr&l=card&f=${encoded}`
      );

      const result = deserializeFromUrl();
      expect(result.queryText).toBe('report');
      expect(result.currentVerticalKey).toBe('documents');
      expect(result.sort).toEqual({ property: 'Created', direction: 'Descending' });
      expect(result.currentPage).toBe(3);
      expect(result.scope).toBe('hr');
      expect(result.activeLayoutKey).toBe('card');
      expect(result.activeFilters).toEqual(filters);
    });

    it('should return empty state for missing parameters', () => {
      setLocationSearch('?sv=1');
      const result = deserializeFromUrl();
      expect(result.queryText).toBeUndefined();
      expect(result.activeFilters).toBeUndefined();
      expect(result.currentVerticalKey).toBeUndefined();
      expect(result.sort).toBeUndefined();
      expect(result.currentPage).toBeUndefined();
      expect(result.scope).toBeUndefined();
      expect(result.activeLayoutKey).toBeUndefined();
    });

    it('should return empty state for completely empty URL', () => {
      setLocationSearch('');
      const result = deserializeFromUrl();
      expect(result).toEqual({});
    });

    describe('with prefix (multi-context)', () => {
      it('should parse prefixed parameters', () => {
        setLocationSearch('?ctx1.sv=1&ctx1.q=budget&ctx1.v=documents');
        const result = deserializeFromUrl('ctx1');
        expect(result.queryText).toBe('budget');
        expect(result.currentVerticalKey).toBe('documents');
      });

      it('should not read non-prefixed params when prefix is specified', () => {
        setLocationSearch('?sv=1&q=unprefixed');
        const result = deserializeFromUrl('ctx1');
        // No ctx1.sv present, so bail early
        expect(result).toEqual({});
      });

      it('should isolate contexts on the same page', () => {
        setLocationSearch('?ctx1.sv=1&ctx1.q=budget&ctx2.sv=1&ctx2.q=people');
        const result1 = deserializeFromUrl('ctx1');
        const result2 = deserializeFromUrl('ctx2');
        expect(result1.queryText).toBe('budget');
        expect(result2.queryText).toBe('people');
      });
    });

    describe('sort parsing edge cases', () => {
      it('should handle sort property containing colons (e.g. custom managed property)', () => {
        // Uses lastIndexOf(':') — property can contain colons
        setLocationSearch('?sv=1&s=ows_MyProp:Descending');
        const result = deserializeFromUrl();
        expect(result.sort).toEqual({
          property: 'ows_MyProp',
          direction: 'Descending',
        });
      });

      it('should ignore sort with no colon', () => {
        setLocationSearch('?sv=1&s=InvalidSortNocolon');
        const result = deserializeFromUrl();
        expect(result.sort).toBeUndefined();
      });
    });
  });

  describe('round-trip serialization/deserialization', () => {
    it('should round-trip a basic state through serialize then deserialize', () => {
      // First, serialize state to URL
      setLocationSearch('');
      const qs = serializeToUrl({
        queryText: 'budget',
        currentVerticalKey: 'documents',
        currentPage: 2,
        activeLayoutKey: 'grid',
      });

      // Then deserialize from that URL
      setLocationSearch('?' + qs);
      const result = deserializeFromUrl();

      expect(result.queryText).toBe('budget');
      expect(result.currentVerticalKey).toBe('documents');
      expect(result.currentPage).toBe(2);
      expect(result.activeLayoutKey).toBe('grid');
    });

    it('should round-trip filters through base64 encoding', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
        { filterName: 'FileType', value: '"pptx"', operator: 'OR' },
        { filterName: 'Author', value: '"Jane Smith"', operator: 'AND' },
      ];

      setLocationSearch('');
      const qs = serializeToUrl({ activeFilters: filters });

      setLocationSearch('?' + qs);
      const result = deserializeFromUrl();
      expect(result.activeFilters).toEqual(filters);
    });

    it('should round-trip sort configuration', () => {
      const sort: ISortField = { property: 'LastModifiedTime', direction: 'Descending' };

      setLocationSearch('');
      const qs = serializeToUrl({ sort });

      setLocationSearch('?' + qs);
      const result = deserializeFromUrl();
      expect(result.sort).toEqual(sort);
    });
  });
});
