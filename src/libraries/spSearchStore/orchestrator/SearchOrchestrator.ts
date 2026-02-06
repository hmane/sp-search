import { StoreApi } from 'zustand/vanilla';
import {
  ISearchStore,
  ISearchQuery,
  ISearchResponse,
  ISearchDataProvider
} from '@interfaces/index';
import { SearchManagerService } from '@services/index';

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
    let prevScope = this._store.getState().scope;
    let prevFilters = this._store.getState().activeFilters;
    let prevVertical = this._store.getState().currentVerticalKey;
    let prevPage = this._store.getState().currentPage;
    let prevSort = this._store.getState().sort;

    this._unsubscribe = this._store.subscribe((state) => {
      const queryChanged = state.queryText !== prevQueryText;
      const scopeChanged = state.scope !== prevScope;
      const filtersChanged = state.activeFilters !== prevFilters;
      const verticalChanged = state.currentVerticalKey !== prevVertical;
      const pageChanged = state.currentPage !== prevPage;
      const sortChanged = state.sort !== prevSort;

      prevQueryText = state.queryText;
      prevScope = state.scope;
      prevFilters = state.activeFilters;
      prevVertical = state.currentVerticalKey;
      prevPage = state.currentPage;
      prevSort = state.sort;

      if (queryChanged || scopeChanged || filtersChanged || verticalChanged || sortChanged) {
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
    this._cancelPending();
  }

  /**
   * Manually trigger a search (e.g., for initial page load).
   */
  public async triggerSearch(): Promise<void> {
    await this._executeSearch();
  }

  private _debouncedSearch(): void {
    if (this._debounceTimer !== undefined) {
      clearTimeout(this._debounceTimer);
    }
    this._debounceTimer = setTimeout(() => {
      this._debounceTimer = undefined;
      this._executeSearch().catch(() => { /* handled in _executeSearch */ });
    }, this._debounceMs);
  }

  private _cancelPending(): void {
    if (this._debounceTimer !== undefined) {
      clearTimeout(this._debounceTimer);
      this._debounceTimer = undefined;
    }
    if (this._abortController) {
      this._abortController.abort();
      this._abortController = undefined;
    }
  }

  private async _executeSearch(): Promise<void> {
    // Cancel any previous in-flight request
    this._cancelPending();

    const state = this._store.getState();
    const provider = this._getProvider(state);

    if (!provider) {
      state.setError('No search data provider registered');
      return;
    }

    // Create new AbortController for this search cycle
    this._abortController = new AbortController();
    const signal = this._abortController.signal;

    // Set loading state
    state.setLoading(true);
    state.setError(undefined);

    try {
      // Build the search query
      const query = this._buildQuery(state);

      // Execute search
      const response: ISearchResponse = await provider.execute(query, signal);

      // Check if aborted during execution
      if (signal.aborted) {
        return;
      }

      // Dispatch results to store
      state.setResults(response.items, response.totalCount);

      // Update refiners if provider supports them
      if (provider.supportsRefiners && response.refiners.length > 0) {
        state.setAvailableRefiners(response.refiners);
      }

      // Update promoted results
      if (response.promotedResults.length > 0) {
        state.setPromotedResults(response.promotedResults);
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

      const message = error instanceof Error ? error.message : 'Search failed';
      state.setError(message);
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

    return {
      queryText: state.queryText || '*',
      queryTemplate: activeVertical?.queryTemplate || state.queryTemplate || '{searchTerms}',
      scope: state.scope,
      filters: state.activeFilters,
      sort: state.sort,
      page: state.currentPage,
      pageSize: state.pageSize,
      selectedProperties: this._getDefaultSelectedProperties(),
      refiners: this._getRefinerProperties(state.filterConfig),
      resultSourceId: activeVertical?.resultSourceId,
      trimDuplicates: true,
    };
  }

  /**
   * Get the primary data provider from the registry.
   */
  private _getProvider(state: ISearchStore): ISearchDataProvider | undefined {
    const providers = state.registries.dataProviders.getAll();
    return providers.length > 0 ? providers[0] : undefined;
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
          queryText: state.queryText || '*',
          queryTemplate: vertical.queryTemplate || state.queryTemplate || '{searchTerms}',
          scope: state.scope,
          filters: state.activeFilters,
          sort: undefined,
          page: 1,
          pageSize: 0,
          selectedProperties: ['Title'],
          refiners: [],
          resultSourceId: vertical.resultSourceId,
          trimDuplicates: true,
        };

        const response = await provider.execute(countQuery, signal);
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
    ];
  }

  /**
   * Extract refiner property names from filter config.
   */
  private _getRefinerProperties(filterConfig: ISearchStore['filterConfig']): string[] {
    if (!filterConfig || filterConfig.length === 0) {
      return [];
    }
    return filterConfig.map((fc) => fc.managedProperty);
  }

  /**
   * Log the search to history (async, non-blocking).
   * Uses full search state for deduplication.
   */
  private _logSearchToHistory(state: ISearchStore, resultCount: number): void {
    if (!this._historyService) {
      return;
    }

    // Build the full search state for hashing and storage
    const searchState = JSON.stringify({
      queryText: state.queryText,
      activeFilters: state.activeFilters,
      currentVerticalKey: state.currentVerticalKey,
      sort: state.sort,
      scope: state.scope,
      activeLayoutKey: state.activeLayoutKey,
    });

    // Log async - don't await, don't block
    this._historyService
      .logSearch(
        state.queryText,
        state.currentVerticalKey,
        state.scope.id,
        searchState,
        resultCount
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
