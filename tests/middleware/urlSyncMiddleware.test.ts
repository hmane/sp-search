import {
  serializeToUrl,
  deserializeFromUrl,
  createUrlSyncSubscription,
} from '../../src/libraries/spSearchStore/store/middleware/urlSyncMiddleware';
import { IActiveFilter, ISortField } from '../../src/libraries/spSearchStore/interfaces';
import { createRegistryContainer } from '../../src/libraries/spSearchStore/registries';
import { createSearchStore } from '../../src/libraries/spSearchStore/store/createStore';

/**
 * Tests for the URL sync middleware — serialization and deserialization
 * of search state to/from URL query parameters.
 *
 * These tests mock window.location to simulate browser behavior in jsdom.
 */
describe('urlSyncMiddleware', () => {
  const originalLocation = window.location;

  function createTestStore() {
    const registries = createRegistryContainer();
    registries.filterTypes.register({
      id: 'checkbox',
      displayName: 'Checkbox',
      component: (() => null) as never,
      serializeValue: (value: unknown): string => String(value),
      deserializeValue: (raw: string): unknown => raw,
      buildRefinementToken: (value: unknown): string => {
        const token = String(value);
        return token.charAt(0) === '"' ? token : '"' + token + '"';
      }
    });
    return createSearchStore(registries);
  }

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

    it('should omit legacy state version tags entirely', () => {
      const result = serializeToUrl({});
      const params = new URLSearchParams(result);
      expect(params.has('x')).toBe(false);
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

    it('should serialize non-default scope as c parameter', () => {
      // 'all' and 'currentsite' are intentionally excluded as common defaults
      // to avoid polluting URLs when the user hasn't changed the scope.
      const result = serializeToUrl({
        scope: { id: 'hub-finance', label: 'Finance Hub' },
      });
      const params = new URLSearchParams(result);
      expect(params.get('c')).toBe('hub-finance');
    });

    it('should omit c for default scopes (all, currentsite)', () => {
      const resultAll = serializeToUrl({ scope: { id: 'all', label: 'All' } });
      expect(new URLSearchParams(resultAll).has('c')).toBe(false);

      const resultSite = serializeToUrl({ scope: { id: 'currentsite', label: 'Current Site' } });
      expect(new URLSearchParams(resultSite).has('c')).toBe(false);
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

    it('should serialize multi-select filters as a single comma-separated alias parameter', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
        { filterName: 'FileType', value: '"pptx"', operator: 'OR' },
      ];
      const result = serializeToUrl({
        activeFilters: filters,
        filterConfig: [
          {
            id: 'f1',
            managedProperty: 'FileType',
            displayName: 'File type',
            urlAlias: 'ft',
            filterType: 'checkbox',
            operator: 'OR',
            multiValues: true,
            maxValues: 10,
            defaultExpanded: true,
            showCount: true,
            sortBy: 'count',
            sortDirection: 'desc',
          },
        ],
      });
      const params = new URLSearchParams(result);
      expect(params.get('ft')).toBe('docx,pptx');
    });

    it('should prefer configured custom aliases over defaults', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'contentclass', value: '"STS_Site"', operator: 'OR' },
      ];
      const result = serializeToUrl({
        activeFilters: filters,
        filterConfig: [
          {
            id: 'f1',
            managedProperty: 'contentclass',
            displayName: 'Content Type',
            urlAlias: 'ct',
            filterType: 'checkbox',
            operator: 'OR',
            multiValues: true,
            maxValues: 10,
            defaultExpanded: true,
            showCount: true,
            sortBy: 'count',
            sortDirection: 'desc',
          },
        ],
      });
      const params = new URLSearchParams(result);
      expect(params.get('ct')).toBe('STS_Site');
      expect(params.has('cc')).toBe(false);
    });

    it('should shorten SharePoint hex refinement tokens in filter URLs', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"ǂǂ646f6378"', operator: 'OR' },
        { filterName: 'FileType', value: '"ǂǂ786c7378"', operator: 'OR' },
      ];
      const result = serializeToUrl({
        activeFilters: filters,
        filterConfig: [
          {
            id: 'f1',
            managedProperty: 'FileType',
            displayName: 'File type',
            urlAlias: 'ft',
            filterType: 'checkbox',
            operator: 'OR',
            multiValues: true,
            maxValues: 10,
            defaultExpanded: true,
            showCount: true,
            sortBy: 'count',
            sortDirection: 'desc',
          },
        ],
      });
      const params = new URLSearchParams(result);
      expect(params.get('ft')).toBe('docx,xlsx');
    });

    it('should safely encode commas inside compact multi-select values', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'ContentType', value: '"Site, Pages"', operator: 'OR' },
        { filterName: 'ContentType', value: '"Knowledge Base"', operator: 'OR' },
      ];
      const result = serializeToUrl({
        activeFilters: filters,
        filterConfig: [
          {
            id: 'f1',
            managedProperty: 'ContentType',
            displayName: 'Content type',
            urlAlias: 'ct',
            filterType: 'checkbox',
            operator: 'OR',
            multiValues: true,
            maxValues: 10,
            defaultExpanded: true,
            showCount: true,
            sortBy: 'count',
            sortDirection: 'desc',
          },
        ],
      });
      const params = new URLSearchParams(result);
      expect(params.get('ct')).toBe('Site%2C%20Pages,Knowledge%20Base');
    });

    it('should omit filter parameters when activeFilters is empty', () => {
      const result = serializeToUrl({ activeFilters: [] });
      const params = new URLSearchParams(result);
      expect(params.has('ft')).toBe(false);
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
      expect(params.get('c')).toBe('hr');
      expect(params.get('l')).toBe('card');
      expect(params.get('ft')).toBe('docx');
    });

    describe('with prefix (multi-context)', () => {
      it('should prefix parameter names', () => {
        const result = serializeToUrl({ queryText: 'test' }, 'ctx1');
        const params = new URLSearchParams(result);
        expect(params.get('ctx1.q')).toBe('test');
      });
    });
  });

  describe('deserializeFromUrl', () => {
    it('should parse state without requiring x tag', () => {
      setLocationSearch('?q=test');
      const result = deserializeFromUrl();
      expect(result.queryText).toBe('test');
      expect(result.urlFilters).toBeUndefined();
    });

    it('should parse queryText from q parameter', () => {
      setLocationSearch('?q=annual+report');
      const result = deserializeFromUrl();
      expect(result.queryText).toBe('annual report');
    });

    it('should parse currentVerticalKey from v parameter', () => {
      setLocationSearch('?v=documents');
      const result = deserializeFromUrl();
      expect(result.currentVerticalKey).toBe('documents');
    });

    it('should parse sort from s parameter (property:direction)', () => {
      setLocationSearch('?s=LastModifiedTime:Descending');
      const result = deserializeFromUrl();
      expect(result.sort).toEqual({
        property: 'LastModifiedTime',
        direction: 'Descending',
      });
    });

    it('should parse sort with Ascending direction', () => {
      setLocationSearch('?s=Title:Ascending');
      const result = deserializeFromUrl();
      expect(result.sort).toEqual({
        property: 'Title',
        direction: 'Ascending',
      });
    });

    it('should ignore sort with invalid direction', () => {
      setLocationSearch('?s=Title:InvalidDir');
      const result = deserializeFromUrl();
      expect(result.sort).toBeUndefined();
    });

    it('should parse currentPage from p parameter', () => {
      setLocationSearch('?p=5');
      const result = deserializeFromUrl();
      expect(result.currentPage).toBe(5);
    });

    it('should ignore non-numeric page values', () => {
      setLocationSearch('?p=abc');
      const result = deserializeFromUrl();
      expect(result.currentPage).toBeUndefined();
    });

    it('should ignore page values less than 1', () => {
      setLocationSearch('?p=0');
      const result = deserializeFromUrl();
      expect(result.currentPage).toBeUndefined();
    });

    it('should parse scope from c parameter', () => {
      setLocationSearch('?c=currentsite');
      const result = deserializeFromUrl();
      expect(result.scope).toBe('currentsite');
    });

    it('should parse activeLayoutKey from l parameter', () => {
      setLocationSearch('?l=grid');
      const result = deserializeFromUrl();
      expect(result.activeLayoutKey).toBe('grid');
    });

    it('should parse compact comma-separated alias parameters', () => {
      setLocationSearch('?ft=docx,pptx&au=John');
      const result = deserializeFromUrl();
      expect(result.urlFilters).toEqual([
        { key: 'ft', rawValue: 'docx,pptx' },
        { key: 'au', rawValue: 'John' },
      ]);
    });

    it('should preserve encoded commas inside compact alias parameters', () => {
      setLocationSearch('?ct=Site%252C%2520Pages,Knowledge%2520Base');
      const result = deserializeFromUrl();
      expect(result.urlFilters).toEqual([
        { key: 'ct', rawValue: 'Site%2C%20Pages,Knowledge%20Base' },
      ]);
    });

    it('should preserve custom alias keys from the URL', () => {
      setLocationSearch('?ct=sts_site');
      const result = deserializeFromUrl();
      expect(result.urlFilters).toEqual([
        { key: 'ct', rawValue: 'sts_site' },
      ]);
    });

    it('should restore legacy base64-encoded f parameters', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"pdf"', operator: 'OR' },
        { filterName: 'Author', value: '"Jane Smith"', operator: 'AND' },
      ];
      const encoded = btoa(unescape(encodeURIComponent(JSON.stringify(filters))));
      setLocationSearch(`?x=1&f=${encoded}`);

      const result = deserializeFromUrl();
      expect(result.activeFilters).toEqual(filters);
      expect(result.urlFilters).toBeUndefined();
    });

    it('should parse a full URL with all parameters', () => {
      setLocationSearch(
        '?q=report&v=documents&s=Created:Descending&p=3&c=hr&l=card&ft=pdf'
      );

      const result = deserializeFromUrl();
      expect(result.queryText).toBe('report');
      expect(result.currentVerticalKey).toBe('documents');
      expect(result.sort).toEqual({ property: 'Created', direction: 'Descending' });
      expect(result.currentPage).toBe(3);
      expect(result.scope).toBe('hr');
      expect(result.activeLayoutKey).toBe('card');
      expect(result.urlFilters).toEqual([
        { key: 'ft', rawValue: 'pdf' },
      ]);
    });

    it('should return empty state for missing parameters', () => {
      setLocationSearch('?');
      const result = deserializeFromUrl();
      expect(result.queryText).toBeUndefined();
      expect(result.urlFilters).toBeUndefined();
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
        setLocationSearch('?ctx1.q=budget&ctx1.v=documents');
        const result = deserializeFromUrl('ctx1');
        expect(result.queryText).toBe('budget');
        expect(result.currentVerticalKey).toBe('documents');
      });

      it('should not read non-prefixed params when prefix is specified', () => {
        setLocationSearch('?q=unprefixed');
        const result = deserializeFromUrl('ctx1');
        expect(result).toEqual({});
      });

      it('should isolate contexts on the same page', () => {
        setLocationSearch('?ctx1.q=budget&ctx2.q=people');
        const result1 = deserializeFromUrl('ctx1');
        const result2 = deserializeFromUrl('ctx2');
        expect(result1.queryText).toBe('budget');
        expect(result2.queryText).toBe('people');
      });
    });

    describe('sort parsing edge cases', () => {
      it('should handle sort property containing colons (e.g. custom managed property)', () => {
        // Uses lastIndexOf(':') — property can contain colons
        setLocationSearch('?s=ows_MyProp:Descending');
        const result = deserializeFromUrl();
        expect(result.sort).toEqual({
          property: 'ows_MyProp',
          direction: 'Descending',
        });
      });

      it('should ignore sort with no colon', () => {
        setLocationSearch('?s=InvalidSortNocolon');
        const result = deserializeFromUrl();
        expect(result.sort).toBeUndefined();
      });

      it('should keep backward compatibility for legacy sc and sid parameters', () => {
        setLocationSearch('?sc=legacy-scope');
        const result = deserializeFromUrl();
        expect(result.scope).toBe('legacy-scope');
      });

      it('should keep backward compatibility for legacy sid parameter', () => {
        setLocationSearch('?sid=25');
        const result = deserializeFromUrl();
        expect(result.stateId).toBe(25);
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

    it('should round-trip filters through compact alias parameters', () => {
      const filters: IActiveFilter[] = [
        { filterName: 'FileType', value: '"docx"', operator: 'OR' },
        { filterName: 'FileType', value: '"pptx"', operator: 'OR' },
        { filterName: 'Author', value: '"Jane Smith"', operator: 'AND' },
      ];

      setLocationSearch('');
      const qs = serializeToUrl({
        activeFilters: filters,
        filterConfig: [
          {
            id: 'f1',
            managedProperty: 'FileType',
            displayName: 'File type',
            urlAlias: 'ft',
            filterType: 'checkbox',
            operator: 'OR',
            multiValues: true,
            maxValues: 10,
            defaultExpanded: true,
            showCount: true,
            sortBy: 'count',
            sortDirection: 'desc',
          },
          {
            id: 'f2',
            managedProperty: 'Author',
            displayName: 'Author',
            urlAlias: 'au',
            filterType: 'people',
            operator: 'AND',
            multiValues: false,
            maxValues: 10,
            defaultExpanded: true,
            showCount: true,
            sortBy: 'count',
            sortDirection: 'desc',
          },
        ],
      });

      setLocationSearch('?' + qs);
      const result = deserializeFromUrl();
      expect(result.urlFilters).toEqual(expect.arrayContaining([
        { key: 'ft', rawValue: 'docx,pptx' },
        { key: 'au', rawValue: 'Jane Smith' },
      ]));
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

  describe('createUrlSyncSubscription', () => {
    it('should replay pending compact filters after filterConfig is registered', () => {
      setLocationSearch('?ft=pdf');
      const store = createTestStore();
      const unsubscribe = createUrlSyncSubscription(store);

      expect(store.getState().activeFilters).toEqual([]);

      store.setState({
        filterConfig: [
          {
            id: 'f1',
            managedProperty: 'FileType',
            displayName: 'File type',
            urlAlias: 'ft',
            filterType: 'checkbox',
            operator: 'OR',
            multiValues: true,
            maxValues: 10,
            defaultExpanded: true,
            showCount: true,
            sortBy: 'count',
            sortDirection: 'desc',
          },
        ],
      });

      expect(store.getState().activeFilters).toEqual([
        { filterName: 'FileType', value: '"pdf"', operator: 'OR', displayValue: undefined },
      ]);

      unsubscribe();
    });
  });
});
