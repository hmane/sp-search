import { StoreApi } from 'zustand/vanilla';
import {
  ISearchStore,
  ISearchQuery,
  ISearchResponse,
  ISearchDataProvider
} from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import { TokenService, ITokenContext } from '@services/TokenService';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { DebugCollector } from '../debug';
import { spLog } from '@store/utils/spLog';

/**
 * SearchOrchestrator — subscribes to store changes and triggers
 * search execution through the registered data provider.
 *
 * Lifecycle:
 *   1. Store query/filter/vertical changes trigger executeSearch()
 *   2. Previous in-flight request is aborted via AbortController
 *   3. Provider.execute() is called with normalized ISearchQuery
 *   4. Results dispatched to resultSlice, refiners to filterSlice
 *   5. Vertical counts fetched in parallel (RowLimit=0 per vertical)
 *   6. Search logged to history (async, non-blocking)
 */
export class SearchOrchestrator {
  private readonly _store: StoreApi<ISearchStore>;
  private _abortController: AbortController | undefined;
  private _unsubscribe: (() => void) | undefined;
  private _debounceTimer: ReturnType<typeof setTimeout> | undefined;
  private readonly _debounceMs: number;
  private _historyService: SearchManagerService | undefined;
  private _lastHistoryId: number = 0;
  /** Tracks whether the registry-freeze guard has run (idempotent). */
  private _registriesFrozen: boolean = false;
  private _firstSearchCompleted: boolean = false;

  public constructor(store: StoreApi<ISearchStore>, debounceMs: number = 300) {
    this._store = store;
    this._debounceMs = debounceMs;
  }

  /**
   * Set the history service for automatic search logging.
   * Optional - if not set, history logging is skipped.
   */
  public setHistoryService(service: SearchManagerService): void {
    this._historyService = service;
  }

  /**
   * Get the last logged history entry ID (for click tracking).
   */
  public getLastHistoryId(): number {
    return this._lastHistoryId;
  }

  /**
   * Start listening to store changes. Call this after providers are registered.
   */
  public start(): void {
    // Track which state fields trigger a new search
    let prevQueryText = this._store.getState().queryText;
    let prevQueryTemplate = this._store.getState().queryTemplate;
    let prevScope = this._store.getState().scope;
    let prevFilters = this._store.getState().activeFilters;
    let prevVertical = this._store.getState().currentVerticalKey;
    let prevPage = this._store.getState().currentPage;
    let prevSort = this._store.getState().sort;
    let prevPageSize = this._store.getState().pageSize;
    let prevSelectedProperties = this._store.getState().selectedProperties;
    let prevResultSourceId = this._store.getState().resultSourceId;
    let prevEnableQueryRules = this._store.getState().enableQueryRules;
    let prevTrimDuplicates = this._store.getState().trimDuplicates;
    let prevCollapseSpecification = this._store.getState().collapseSpecification;
    let prevRefinementFilters = this._store.getState().refinementFilters;
    // Use JSON for filterConfig comparison — React components can re-set the
    // same filterConfig array (new reference, same content) on mount, which
    // triggers a false-positive change detection and aborts in-flight searches.
    let prevFilterConfigJson = JSON.stringify(this._store.getState().filterConfig);
    let prevOperatorBetweenFilters = this._store.getState().operatorBetweenFilters;
    let prevQueryInputTransformation = this._store.getState().queryInputTransformation;

    this._unsubscribe = this._store.subscribe((state) => {
      const queryChanged = state.queryText !== prevQueryText;
      const queryTemplateChanged = state.queryTemplate !== prevQueryTemplate;
      const scopeChanged = state.scope !== prevScope;
      const filtersChanged = state.activeFilters !== prevFilters;
      const verticalChanged = state.currentVerticalKey !== prevVertical;
      const pageChanged = state.currentPage !== prevPage;
      const sortChanged = state.sort !== prevSort;
      const pageSizeChanged = state.pageSize !== prevPageSize;
      const selectedPropertiesChanged = state.selectedProperties !== prevSelectedProperties;
      const resultSourceChanged = state.resultSourceId !== prevResultSourceId;
      const queryRulesChanged = state.enableQueryRules !== prevEnableQueryRules;
      const trimDuplicatesChanged = state.trimDuplicates !== prevTrimDuplicates;
      const collapseChanged = state.collapseSpecification !== prevCollapseSpecification;
      const refinementFiltersChanged = state.refinementFilters !== prevRefinementFilters;
      const currentFilterConfigJson = JSON.stringify(state.filterConfig);
      const filterConfigChanged = currentFilterConfigJson !== prevFilterConfigJson;
      const operatorChanged = state.operatorBetweenFilters !== prevOperatorBetweenFilters;
      const transformationChanged = state.queryInputTransformation !== prevQueryInputTransformation;

      prevQueryText = state.queryText;
      prevQueryTemplate = state.queryTemplate;
      prevScope = state.scope;
      prevFilters = state.activeFilters;
      prevVertical = state.currentVerticalKey;
      prevPage = state.currentPage;
      prevSort = state.sort;
      prevPageSize = state.pageSize;
      prevSelectedProperties = state.selectedProperties;
      prevResultSourceId = state.resultSourceId;
      prevEnableQueryRules = state.enableQueryRules;
      prevTrimDuplicates = state.trimDuplicates;
      prevCollapseSpecification = state.collapseSpecification;
      prevRefinementFilters = state.refinementFilters;
      prevFilterConfigJson = currentFilterConfigJson;
      prevOperatorBetweenFilters = state.operatorBetweenFilters;
      prevQueryInputTransformation = state.queryInputTransformation;

      // Auto-switch to the vertical's configured defaultLayout when the vertical changes.
      if (verticalChanged && state.currentVerticalKey) {
        const newVertical = state.verticals.find((v) => v.key === state.currentVerticalKey);
        const targetLayout = newVertical?.defaultLayout;
        if (targetLayout &&
            state.availableLayouts.indexOf(targetLayout) >= 0 &&
            state.activeLayoutKey !== targetLayout) {
          // setTimeout(0) is REQUIRED here — do not change to queueMicrotask().
          // setLayout() triggers a Zustand setState inside a subscription callback.
          // setTimeout defers to the next macro task, breaking the re-entry cycle.
          // queueMicrotask would run within the same subscription call stack, causing
          // infinite re-entry. The one-frame flicker is an acceptable trade-off.
          setTimeout((): void => {
            this._store.getState().setLayout(targetLayout);
          }, 0);
        }
      }

      if (
        queryChanged ||
        queryTemplateChanged ||
        scopeChanged ||
        filtersChanged ||
        verticalChanged ||
        sortChanged ||
        pageSizeChanged ||
        selectedPropertiesChanged ||
        resultSourceChanged ||
        queryRulesChanged ||
        trimDuplicatesChanged ||
        collapseChanged ||
        refinementFiltersChanged ||
        filterConfigChanged ||
        operatorChanged ||
        transformationChanged
      ) {
        // Reset to page 1 for non-page changes (page is already set by the slice)
        this._debouncedSearch();
      } else if (pageChanged) {
        // Page changes execute immediately (no debounce)
        this._executeSearch().catch(() => { /* handled in _executeSearch */ });
      }
    });
  }

  /**
   * Stop listening and abort any in-flight request.
   */
  public stop(): void {
    if (this._unsubscribe) {
      this._unsubscribe();
      this._unsubscribe = undefined;
    }
    this.cancelPending();
  }

  /**
   * Manually trigger a search (e.g., for initial page load).
   */
  public async triggerSearch(): Promise<void> {
    await this._executeSearch();
  }

  /**
   * Freeze all registries to prevent mid-session mutations.
   * Call this from the Results web part after ALL providers, actions,
   * layouts, and filter types have been registered. This must NOT be
   * called automatically from _executeSearch because web parts load
   * asynchronously and the first search can fire before all web parts
   * have finished registering their providers.
   *
   * Note: suggestions registry is NOT frozen because suggestion
   * providers register after initializeSearchContext.
   */
  public freezeRegistries(): void {
    if (this._registriesFrozen) { return; }
    this._registriesFrozen = true;
    const r = this._store.getState().registries;
    r.dataProviders.freeze();
    r.actions.freeze();
    r.layouts.freeze();
    r.filterTypes.freeze();
  }

  private _debouncedSearch(): void {
    if (this._debounceTimer !== undefined) {
      clearTimeout(this._debounceTimer);
    }
    // _executeSearch() reads fresh state via this._store.getState()
    // at call time, so state changes during debounce window are captured.
    this._debounceTimer = setTimeout(() => {
      this._debounceTimer = undefined;
      this._executeSearch().catch(() => { /* handled in _executeSearch */ });
    }, this._debounceMs);
  }

  /**
   * Cancel any pending debounce timer + abort the in-flight provider call.
   * Safe to call multiple times; safe to call when nothing is pending.
   * Public so admin/diagnostic UI can wire a "Cancel" button via
   * `getOrchestrator(searchContextId).cancelPending()`.
   */
  public cancelPending(): void {
    if (this._debounceTimer !== undefined) {
      clearTimeout(this._debounceTimer);
      this._debounceTimer = undefined;
    }
    if (this._abortController) {
      this._abortController.abort();
      this._abortController = undefined;
    }
  }

  /**
   * Execute a provider search with automatic QuotaExceededError retry.
   * PnPjs caching middleware may throw when storage is full during its
   * cache-write step, even though the API returned valid data. This
   * aggressively cleans both localStorage and sessionStorage then retries once.
   *
   * T5.D2 — emits a `logNetworkEvent` entry per call (including the
   * retry-after-quota path) so the DebugPanel's Network tab can render
   * timing badges. `kind` distinguishes the main search from vertical-
   * count fan-out; `verticalKey` is set for vertical counts only.
   */
  private async _executeProviderWithRetry(
    provider: ISearchDataProvider,
    query: ISearchQuery,
    signal: AbortSignal,
    kind: 'search' | 'verticalCount' = 'search',
    verticalKey?: string
  ): Promise<ISearchResponse> {
    const start = performance.now();
    const baseEntry = {
      providerId: provider.id,
      kind,
      queryTemplate: query.queryTemplate,
      currentPage: query.page,
      pageSize: query.pageSize,
      verticalKey,
    };
    try {
      const response = await provider.execute(query, signal);
      DebugCollector.logNetworkEvent({
        ...baseEntry,
        status: 'ok',
        durationMs: Math.round(performance.now() - start),
        totalCount: response.totalCount,
        itemCount: response.items.length,
        errorMessage: undefined,
      });
      return response;
    } catch (error) {
      if (error instanceof DOMException && error.name === 'QuotaExceededError') {
        // Retry once — space should now be available.
        this._cleanupStorage();
        const retryStart = performance.now();
        try {
          const retryResponse = await provider.execute(query, signal);
          DebugCollector.logNetworkEvent({
            ...baseEntry,
            status: 'ok',
            durationMs: Math.round(performance.now() - retryStart),
            totalCount: retryResponse.totalCount,
            itemCount: retryResponse.items.length,
            errorMessage: 'recovered after QuotaExceededError retry',
          });
          return retryResponse;
        } catch (retryError) {
          DebugCollector.logNetworkEvent({
            ...baseEntry,
            status: signal.aborted ? 'aborted' : 'error',
            durationMs: Math.round(performance.now() - retryStart),
            totalCount: undefined,
            itemCount: undefined,
            errorMessage: retryError instanceof Error ? retryError.message : String(retryError),
          });
          throw retryError;
        }
      }
      DebugCollector.logNetworkEvent({
        ...baseEntry,
        status: error instanceof DOMException && error.name === 'AbortError' ? 'aborted' : 'error',
        durationMs: Math.round(performance.now() - start),
        totalCount: undefined,
        itemCount: undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
      });
      throw error;
    }
  }

  /**
   * Aggressively clean both localStorage and sessionStorage to free space.
   * Removes: SPFx numeric hash keys, PnP cache entries (pnp-*).
   */
  private _cleanupStorage(): void {
    try {
      this._cleanupStorageInstance(localStorage);
    } catch { /* swallow */ }
    try {
      this._cleanupStorageInstance(sessionStorage);
    } catch { /* swallow */ }
  }

  private _cleanupStorageInstance(storage: Storage): void {
    const keysToRemove: string[] = [];
    for (let i = 0; i < storage.length; i++) {
      const key = storage.key(i);
      if (key && (/^-?\d+$/.test(key) || key.indexOf('pnp-') === 0)) {
        keysToRemove.push(key);
      }
    }
    for (let i = 0; i < keysToRemove.length; i++) {
      storage.removeItem(keysToRemove[i]);
    }
  }

  private async _executeSearch(): Promise<void> {
    // Cancel any previous in-flight request
    this.cancelPending();

    // Freeze the provider/action/layout/filterType registries on the first
    // search execution. By this point every web part's onInit has completed
    // (SPFx runs them all before yielding) and the 300ms debounce window
    // has elapsed, so any late-arriving Filters/Verticals registrations
    // have already landed. Eliminates the race where Results.onInit froze
    // registries before Filters.onInit registered its built-in filter types.
    this.freezeRegistries();

    const state = this._store.getState();
    const provider = this._getProvider(state);

    if (!provider) {
      // No provider yet — this is normal when the Filters or Verticals web part
      // initializes before the Results web part registers the data provider.
      // Skip silently; the Results web part will trigger a search after registering.
      spLog.warn('Search skipped — no data provider registered. Register a provider (e.g. SharePointSearchProvider) before calling triggerSearch().');
      return;
    }

    if (!this._firstSearchCompleted) {
      // T5.D6 — `queryText` is auto-redacted by spLog so the F12 console
      // never sees a literal user query, even in debug mode.
      spLog.info('First search starting', {
        providerId: provider.id,
        queryText: state.queryText,
        activeFilterCount: state.activeFilters.length,
        currentPage: state.currentPage,
      });
    }

    // Create new AbortController for this search cycle
    this._abortController = new AbortController();
    const signal = this._abortController.signal;
    const searchStart = performance.now();

    // Set loading state
    state.setLoading(true);
    state.setError(undefined);

    try {
      // Build the search query
      const query = this._buildQuery(state);

      // Debug: capture query info before execution
      DebugCollector.setLastQuery({
        kql: query.queryText,
        queryTemplate: query.queryTemplate,
        resultSourceId: query.resultSourceId,
        refinementFilters: query.filters.map((f) => f.filterName + ':' + f.value),
        providerId: provider.id,
        startTime: searchStart,
        duration: undefined,
        totalCount: undefined,
        itemsReturned: undefined,
        currentPage: query.page,
        pageSize: query.pageSize,
        refiners: [],
        error: undefined,
        request: query as unknown as Record<string, unknown>,
        response: undefined,
      });

      // Detect first search after URL restore with active filters.
      // displayRefiners is empty on page load, so the merge in filterSlice has
      // nothing to merge with. Fire a parallel lightweight search WITHOUT filters
      // to get the full refiner set — this populates the base before the filtered
      // refiners merge in.
      const needBaseRefiners = provider.supportsRefiners
        && state.activeFilters.length > 0
        && state.displayRefiners.length === 0
        && query.refiners.length > 0;

      // Execute main search + optional base refiner query in parallel.
      // Wrapped with QuotaExceededError retry — PnPjs may throw when
      // localStorage is full during its caching step.
      const mainPromise: Promise<ISearchResponse> = this._executeProviderWithRetry(provider, query, signal);
      const basePromise: Promise<ISearchResponse | undefined> = needBaseRefiners
        ? this._executeProviderWithRetry(provider, {
            queryText: query.queryText,
            queryTemplate: query.queryTemplate,
            scope: query.scope,
            filters: [],             // No filters — get full refiner set
            sort: undefined,
            page: 1,
            pageSize: 0,             // No results needed — we only need refiners
            selectedProperties: ['Title'],
            refiners: query.refiners,
            resultSourceId: query.resultSourceId,
            trimDuplicates: true,
            refinementFilters: query.refinementFilters,
          }, signal).catch(function (): undefined { return undefined; })
        : Promise.resolve(undefined);

      const [response, baseResponse] = await Promise.all([mainPromise, basePromise]);

      // Check if aborted during execution
      if (signal.aborted) {
        return;
      }

      // Dispatch results to store
      // SharePoint TotalRows is an ESTIMATE — it can overcount when
      // TrimDuplicates, CollapseSpecification, or security trimming reduce
      // actual results. If a page returns 0 items but we're past page 1,
      // cap totalCount to prevent navigating to empty pages.
      let adjustedTotal = response.totalCount;
      if (response.items.length === 0 && state.currentPage > 1) {
        adjustedTotal = (state.currentPage - 1) * state.pageSize;
        // Reset to the last valid page
        state.setPage(state.currentPage - 1);
      }
      state.setResults(response.items, adjustedTotal);

      // Update refiners if provider supports them
      if (provider.supportsRefiners) {
        if (baseResponse && baseResponse.refiners.length > 0) {
          // Set base (unfiltered) refiners first — populates displayRefiners
          state.setAvailableRefiners(baseResponse.refiners);
          // Now merge filtered refiners in — updates counts for matching values,
          // keeps all other values visible with count 0
          if (response.refiners.length > 0) {
            state.setAvailableRefiners(response.refiners);
          }
        } else if (response.refiners.length > 0) {
          state.setAvailableRefiners(response.refiners);
        }
      }

      // Update promoted results
      if (response.promotedResults.length > 0) {
        state.setPromotedResults(response.promotedResults);
      }

      // Update "Did you mean" suggestion from search API
      state.setQuerySuggestion(response.querySuggestion || undefined);

      // All synchronous result processing is done — clear loading state
      state.setLoading(false);

      const elapsed = Math.round(performance.now() - searchStart);

      // Debug: update query info with results + log SEARCH event
      DebugCollector.setLastQuery({
        kql: query.queryText,
        queryTemplate: query.queryTemplate,
        resultSourceId: query.resultSourceId,
        refinementFilters: query.filters.map((f) => f.filterName + ':' + f.value),
        providerId: provider.id,
        startTime: searchStart,
        duration: elapsed,
        totalCount: adjustedTotal,
        itemsReturned: response.items.length,
        currentPage: query.page,
        pageSize: query.pageSize,
        refiners: response.refiners.map((r) => ({
          name: r.filterName,
          values: r.values.map((v) => ({ value: v.value, count: v.count })),
        })),
        error: undefined,
        request: query as unknown as Record<string, unknown>,
        response: {
          totalCount: adjustedTotal,
          itemCount: response.items.length,
          refinersCount: response.refiners.length,
          promotedResultsCount: response.promotedResults.length,
          querySuggestion: response.querySuggestion || undefined,
          items: response.items.slice(0, 5).map((item) => ({
            title: item.title,
            url: item.url,
            fileType: item.fileType,
          })),
        },
      });
      DebugCollector.logEvent('SEARCH', {
        duration: elapsed,
        resultCount: response.items.length,
        totalCount: adjustedTotal,
        providerId: provider.id,
        query: query.queryText,
      });

      if (!this._firstSearchCompleted) {
        this._firstSearchCompleted = true;
        spLog.info('First search complete', {
          resultCount: response.items.length,
          totalCount: adjustedTotal,
          durationMs: elapsed,
          providerId: provider.id,
        });
      } else {
        spLog.debug('Search complete', {
          resultCount: response.items.length,
          totalCount: adjustedTotal,
          durationMs: elapsed,
          queryText: state.queryText,
        });
      }

      // Fetch vertical counts in parallel
      this._fetchVerticalCounts(provider, state, signal);

      // Log search to history (async, non-blocking)
      this._logSearchToHistory(state, response.totalCount);

    } catch (error) {
      // Don't show error for user-cancelled requests
      if (error instanceof DOMException && error.name === 'AbortError') {
        return;
      }

      if (signal.aborted) {
        return;
      }

      // Swallow storage quota errors — these come from SPFx framework
      // serializing web part properties to localStorage, not from search.
      if (error instanceof DOMException && error.name === 'QuotaExceededError') {
        state.setLoading(false);
        return;
      }

      spLog.error('Search failed', {
        providerId: provider.id,
        queryText: state.queryText,
        errorMessage: error instanceof Error ? error.message : String(error),
      });
      const message = error instanceof Error ? error.message : 'Search failed';
      state.setError(message);
      DebugCollector.logEvent('ERROR', {
        message,
        providerId: provider.id,
        query: state.queryText || '*',
        stack: error instanceof Error ? error.stack : undefined,
      });
      state.setLoading(false);
    }
  }

  /**
   * Build normalized ISearchQuery from current store state.
   */
  private _buildQuery(state: ISearchStore): ISearchQuery {
    // Get the active vertical's configuration
    const activeVertical = state.verticals.find(
      (v) => v.key === state.currentVerticalKey
    );

    const splitFilters = this._splitActiveFilters(state);
    const queryText = this._buildEffectiveQueryText(state, splitFilters.textFilters);

    return {
      queryText,
      queryTemplate: activeVertical?.queryTemplate || state.queryTemplate || '{searchTerms}',
      scope: state.scope,
      filters: splitFilters.refinementFilters,
      operatorBetweenFilters: state.operatorBetweenFilters,
      sort: state.sort,
      page: state.currentPage,
      pageSize: state.pageSize,
      selectedProperties: this._mergeSelectedProperties(state.selectedProperties),
      refiners: this._getRefinerProperties(state.filterConfig),
      resultSourceId: activeVertical?.resultSourceId || state.resultSourceId || undefined,
      trimDuplicates: state.trimDuplicates,
      enableQueryRules: state.enableQueryRules,
      collapseSpecification: state.collapseSpecification || undefined,
      refinementFilters: state.refinementFilters || undefined,
    };
  }

  /**
   * Get the data provider for the current vertical.
   * If the active vertical specifies a dataProviderId, that provider is used.
   * Falls back to the first registered provider.
   */
  private _getProvider(state: ISearchStore): ISearchDataProvider | undefined {
    const providers = state.registries.dataProviders.getAll();
    if (providers.length === 0) {
      return undefined;
    }
    const activeVertical = state.verticals.find((v) => v.key === state.currentVerticalKey);
    if (activeVertical?.dataProviderId) {
      const specific = providers.find((p) => p.id === activeVertical.dataProviderId);
      if (specific) {
        return specific;
      }
    }
    return providers[0];
  }

  /**
   * Fetch counts for all verticals in parallel.
   * Uses RowLimit=0 to get just the count without results.
   */
  private _fetchVerticalCounts(
    provider: ISearchDataProvider,
    state: ISearchStore,
    signal: AbortSignal
  ): void {
    const verticals = state.verticals;
    if (verticals.length <= 1) {
      return;
    }

    const countPromises = verticals.map(async (vertical) => {
      try {
        const countQuery: ISearchQuery = {
          queryText: this._buildEffectiveQueryText(state, this._splitActiveFilters(state).textFilters),
          queryTemplate: vertical.queryTemplate || state.queryTemplate || '{searchTerms}',
          scope: state.scope,
          filters: this._splitActiveFilters(state).refinementFilters,
          sort: undefined,
          page: 1,
          pageSize: 0,
          selectedProperties: ['Title'],
          refiners: [],
          resultSourceId: vertical.resultSourceId,
          trimDuplicates: true,
        };

        // T5.D2 — label vertical-count traffic so the DebugPanel Network
        // tab distinguishes it from the main search request.
        const response = await this._executeProviderWithRetry(provider, countQuery, signal, 'verticalCount', vertical.key);
        return { key: vertical.key, count: response.totalCount };
      } catch {
        return { key: vertical.key, count: 0 };
      }
    });

    Promise.all(countPromises)
      .then((counts) => {
        if (signal.aborted) {
          return;
        }
        const countMap: Record<string, number> = {};
        for (const c of counts) {
          countMap[c.key] = c.count;
        }
        state.setVerticalCounts(countMap);
      })
      .catch(() => {
        // Swallow errors for count queries
      });
  }

  /**
   * Default managed properties to retrieve.
   */
  private _getDefaultSelectedProperties(): string[] {
    return [
      'Title', 'Path', 'Filename', 'Author', 'AuthorOWSUSER',
      'Created', 'LastModifiedTime', 'FileType', 'FileExtension',
      'SecondaryFileExtension', 'contentclass',
      'HitHighlightedSummary', 'HitHighlightedProperties',
      'SiteName', 'SiteTitle', 'SPSiteURL',
      'ServerRedirectedURL', 'ServerRedirectedPreviewURL',
      'PictureThumbnailURL', 'ParentLink', 'ViewsLifeTime',
      'Size', 'NormSiteID', 'NormListID', 'NormUniqueID',
      'DocId', 'IsDocument', 'UniqueId',
      'ListId', 'ListItemID',
    ];
  }

  /**
   * Merge admin-configured selected properties with the default set.
   * Custom properties are appended; duplicates are deduplicated.
   */
  private _mergeSelectedProperties(customProperties: string): string[] {
    const defaults = this._getDefaultSelectedProperties();
    if (!customProperties) {
      return defaults;
    }
    const custom = customProperties.split(',').map((p) => p.trim()).filter(Boolean);
    if (custom.length === 0) {
      return defaults;
    }
    const seen = new Set(defaults);
    for (const prop of custom) {
      if (!seen.has(prop)) {
        defaults.push(prop);
        seen.add(prop);
      }
    }
    return defaults;
  }

  /**
   * Extract refiner property names from filter config.
   */
  private _getRefinerProperties(filterConfig: ISearchStore['filterConfig']): string[] {
    if (!filterConfig || filterConfig.length === 0) {
      return [];
    }

    return filterConfig
      .filter((fc) => this._usesBucketedRefiners(fc.filterType))
      .map((fc) => fc.managedProperty);
  }

  private _usesBucketedRefiners(filterType: ISearchStore['filterConfig'][number]['filterType']): boolean {
    return filterType === 'checkbox' ||
      filterType === 'dropdown' ||
      filterType === 'tagbox' ||
      filterType === 'slider' ||
      filterType === 'taxonomy';
  }

  private _splitActiveFilters(state: ISearchStore): {
    textFilters: ISearchStore['activeFilters'];
    refinementFilters: ISearchStore['activeFilters'];
  } {
    const textFilters: ISearchStore['activeFilters'] = [];
    const refinementFilters: ISearchStore['activeFilters'] = [];
    const configMap = new Map<string, ISearchStore['filterConfig'][number]>();

    for (let i = 0; i < state.filterConfig.length; i++) {
      configMap.set(state.filterConfig[i].managedProperty, state.filterConfig[i]);
    }

    for (let i = 0; i < state.activeFilters.length; i++) {
      const filter = state.activeFilters[i];
      const config = configMap.get(filter.filterName);
      if (config?.filterType === 'text') {
        textFilters.push(filter);
      } else {
        refinementFilters.push(filter);
      }
    }

    return { textFilters, refinementFilters };
  }

  /**
   * MISS-001 — build an ITokenContext from SPContext for resolving tokens
   * other than `{searchTerms}` in admin-configured templates. Returns
   * empty strings for fields when SPContext isn't ready (e.g. during the
   * very first render before `onInit` resolves) — TokenService treats
   * empty replacements as no-ops, so this degrades gracefully.
   */
  private _buildTokenContext(rawQuery: string): ITokenContext {
    const pageContext = SPContext.isReady() ? SPContext.pageContext : undefined;
    const user = pageContext?.user;
    return {
      queryText: rawQuery || '',
      siteId: pageContext?.site?.id?.toString() || '',
      siteUrl: pageContext?.site?.absoluteUrl || '',
      webId: pageContext?.web?.id?.toString() || '',
      webUrl: pageContext?.web?.absoluteUrl || '',
      hubSiteUrl: (pageContext?.legacyPageContext as { hubSiteId?: string; hubSiteUrl?: string })?.hubSiteUrl || '',
      userDisplayName: user?.displayName || '',
      userEmail: user?.email || user?.loginName || '',
      listId: pageContext?.list?.id?.toString() || '',
    };
  }

  private _buildEffectiveQueryText(
    state: ISearchStore,
    textFilters: ISearchStore['activeFilters']
  ): string {
    const rawQuery = state.queryText;
    const transformation = state.queryInputTransformation || '{searchTerms}';
    // MISS-001 — full token resolution (not just `{searchTerms}`). Admin
    // patterns like `({searchTerms}) AND owner:{User.Email}` now expand
    // every token client-side before the provider sees the query.
    const tokenContext = this._buildTokenContext(rawQuery);
    const queryText = TokenService.applyQueryInputTransformation(
      transformation,
      rawQuery,
      tokenContext
    );

    if (textFilters.length === 0) {
      return queryText;
    }

    const configMap = new Map<string, ISearchStore['filterConfig'][number]>();
    for (let i = 0; i < state.filterConfig.length; i++) {
      configMap.set(state.filterConfig[i].managedProperty, state.filterConfig[i]);
    }

    const textClauses: string[] = [];
    for (let i = 0; i < textFilters.length; i++) {
      const filter = textFilters[i];
      const config = configMap.get(filter.filterName);
      const clause = this._buildTextFilterClause(filter.filterName, filter.value, config?.operator || 'AND');
      if (clause) {
        textClauses.push(clause);
      }
    }

    if (textClauses.length === 0) {
      return queryText;
    }

    if (queryText === '*') {
      return textClauses.join(' AND ');
    }

    return '(' + queryText + ') AND ' + textClauses.join(' AND ');
  }

  private _buildTextFilterClause(
    managedProperty: string,
    rawValue: string,
    operator: 'AND' | 'OR'
  ): string | undefined {
    const terms = rawValue
      .trim()
      .split(/\s+/)
      .map((term) => this._escapeKqlTerm(term))
      .filter(Boolean);

    if (terms.length === 0) {
      return undefined;
    }

    if (terms.length === 1) {
      return managedProperty + ':' + terms[0];
    }

    return managedProperty + ':(' + terms.join(' ' + operator + ' ') + ')';
  }

  private _escapeKqlTerm(value: string): string {
    const escaped = value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    return '"' + escaped + '"';
  }

  /**
   * Log the search to history (async, non-blocking).
   * Uses full search state for deduplication.
   */
  private _logSearchToHistory(state: ISearchStore, resultCount: number): void {
    if (!this._historyService) {
      return;
    }

    const hasQueryText = Boolean(state.queryText && state.queryText.trim());
    const hasActiveFilters = state.activeFilters.length > 0;

    // Skip passive browse loads. These are auto-triggered hydration searches,
    // not user-driven search history entries.
    if (!hasQueryText && !hasActiveFilters) {
      return;
    }

    // Build the full search state for hashing and storage
    const searchState = JSON.stringify({
      queryText: state.queryText,
      activeFilters: state.activeFilters,
      currentVerticalKey: state.currentVerticalKey,
      sort: state.sort,
    });

    const searchPageUrl = typeof window !== 'undefined' ? window.location.pathname : '';

    // Log async - don't await, don't block
    this._historyService
      .logSearch(
        state.queryText,
        state.currentVerticalKey,
        searchPageUrl,
        searchState,
        resultCount,
        resultCount === 0
      )
      .then((historyId) => {
        this._lastHistoryId = historyId;
      })
      .catch(() => {
        // Non-critical - swallow errors
      });
  }

  /**
   * Log a clicked item to the current history entry.
   */
  public logClickedItem(clickedUrl: string, clickedTitle: string, position: number): void {
    if (!this._historyService || this._lastHistoryId <= 0) {
      return;
    }

    // Log async - don't await, don't block
    this._historyService
      .logClickedItem(this._lastHistoryId, clickedUrl, clickedTitle, position)
      .catch(() => {
        // Non-critical - swallow errors
      });
  }
}
