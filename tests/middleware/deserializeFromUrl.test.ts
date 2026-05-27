import { deserializeFromUrl } from '../../src/libraries/spSearchStore/store/middleware/urlSyncMiddleware';

/**
 * Focused regression tests for URL → state deserialization, specifically the
 * reserved-param exclusion that keeps SharePoint / SPFx / DebugFab system
 * params (`debug`, `noredir`, `loadSPFX`, `debugManifestsFile`, …) from being
 * mistaken for filter aliases — which previously made the URL-sync middleware
 * wait `URL_FILTER_RESTORE_TIMEOUT_MS` (5 s) for a `filterConfig` that never
 * arrives, on every page loaded with `?debug=1`.
 *
 * (The broader `urlSyncMiddleware.test.ts` suite is deferred per
 * `config/jest.config.README.md` — this file is intentionally separate and
 * only exercises the exported `deserializeFromUrl` surface.)
 */
describe('deserializeFromUrl — reserved-param exclusion', () => {
  let originalLocation: Location;

  function setSearch(search: string): void {
    Object.defineProperty(window, 'location', {
      writable: true,
      configurable: true,
      value: { ...window.location, search },
    });
  }

  beforeEach(() => {
    originalLocation = window.location;
  });

  afterEach(() => {
    Object.defineProperty(window, 'location', {
      writable: true,
      configurable: true,
      value: originalLocation,
    });
  });

  it('does not treat ?debug=1 as a filter param', () => {
    setSearch('?debug=1');
    const state = deserializeFromUrl();
    expect(state.urlFilters).toBeUndefined();
  });

  it('does not treat the SPFx debug-serve params as filter params', () => {
    setSearch('?debug=true&noredir=true&loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/build/manifests.js');
    const state = deserializeFromUrl();
    expect(state.urlFilters).toBeUndefined();
  });

  it('still treats unknown non-reserved params as filter params', () => {
    setSearch('?ft=docx,pptx');
    const state = deserializeFromUrl();
    expect(state.urlFilters).toEqual([{ key: 'ft', rawValue: 'docx,pptx' }]);
  });

  it('keeps filter params while dropping the debug param when both are present', () => {
    setSearch('?debug=1&ft=docx');
    const state = deserializeFromUrl();
    expect(state.urlFilters).toEqual([{ key: 'ft', rawValue: 'docx' }]);
  });

  it('still parses the documented short reserved params into state, not urlFilters', () => {
    setSearch('?q=hello&v=docs&s=Title:Ascending&p=2&l=grid');
    const state = deserializeFromUrl();
    expect(state.queryText).toBe('hello');
    expect(state.currentVerticalKey).toBe('docs');
    expect(state.sort).toEqual({ property: 'Title', direction: 'Ascending' });
    expect(state.currentPage).toBe(2);
    expect(state.activeLayoutKey).toBe('grid');
    expect(state.urlFilters).toBeUndefined();
  });
});
