import 'spfx-toolkit/lib/utilities/context/pnpImports/lists';
import 'spfx-toolkit/lib/utilities/context/pnpImports/security';
import { CacheNever } from '@pnp/queryable';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { createSPExtractor } from 'spfx-toolkit/lib/utilities/listItemHelper';
import { BatchBuilder } from 'spfx-toolkit/lib/utilities/batchBuilder';
import {
  ISavedSearch,
  ISearchCollection,
  ICollectionItem,
  ISearchHistoryEntry,
  IPersonaInfo
} from '@interfaces/index';

// ─── List Names ────────────────────────────────────────────
const SAVED_QUERIES_LIST = 'SearchSavedQueries';
const HISTORY_LIST = 'SearchHistory';
const COLLECTIONS_LIST = 'SearchCollections';
const HISTORY_RETENTION_DAYS = 90;
const HISTORY_CLEANUP_INTERVAL_MS = 24 * 60 * 60 * 1000;

interface IHistoryStateSummary {
  activeFilters?: Array<{
    filterName?: string;
    displayValue?: string;
    value?: string;
  }>;
  currentVerticalKey?: string;
}

/**
 * Compute SHA-256 hash of the full search state for deduplication.
 * Includes queryText, filters, vertical, scope, and sort.
 */
async function computeStateHash(stateJson: string): Promise<string> {
  if (typeof crypto !== 'undefined' && crypto.subtle) {
    const data = new TextEncoder().encode(stateJson);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map((b) => ('00' + b.toString(16)).slice(-2)).join('');
  }
  // Fallback: simple hash for environments without crypto.subtle
  let hash = 0;
  for (let i = 0; i < stateJson.length; i++) {
    const char = stateJson.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return Math.abs(hash).toString(16);
}

/**
 * Extract IPersonaInfo[] from a multi-value person field using SPExtractor.
 */
function extractSharedWith(item: Record<string, unknown>): IPersonaInfo[] {
  const sharedWith: IPersonaInfo[] = [];
  const rawShared = item.SharedWith as Array<Record<string, unknown>> | undefined;
  if (rawShared && Array.isArray(rawShared)) {
    for (let i = 0; i < rawShared.length; i++) {
      const userExt = createSPExtractor(rawShared[i]);
      sharedWith.push({
        displayText: userExt.string('Title', ''),
        email: userExt.string('EMail', ''),
        imageUrl: undefined,
      });
    }
  }
  return sharedWith;
}

/**
 * Extract author IPersonaInfo from Author expand field.
 */
function extractAuthor(item: Record<string, unknown>): IPersonaInfo {
  const authorRaw = item.Author as Record<string, unknown> | undefined;
  if (authorRaw) {
    const ext = createSPExtractor(authorRaw);
    return {
      displayText: ext.string('Title', ''),
      email: ext.string('EMail', ''),
      imageUrl: undefined,
    };
  }
  return { displayText: '', email: '', imageUrl: undefined };
}

/**
 * Maps a raw SharePoint list item to ISavedSearch using SPExtractor.
 */
function mapToSavedSearch(item: Record<string, unknown>): ISavedSearch {
  const ext = createSPExtractor(item);

  return {
    id: ext.number('Id', 0),
    title: ext.string('Title', ''),
    queryText: ext.string('QueryText', ''),
    searchState: ext.string('SearchState', '{}'),
    searchUrl: ext.string('SearchUrl', ''),
    entryType: (ext.string('EntryType', 'SavedSearch') as 'SavedSearch' | 'SharedSearch'),
    category: ext.string('Category', ''),
    sharedWith: extractSharedWith(item),
    resultCount: ext.number('ResultCount', 0),
    lastUsed: ext.date('LastUsed') || ext.date('Created') || new Date(),
    created: ext.date('Created') || new Date(),
    author: extractAuthor(item),
  };
}

/**
 * Maps a raw SharePoint list item to ISearchHistoryEntry using SPExtractor.
 */
function mapToHistoryEntry(item: Record<string, unknown>): ISearchHistoryEntry {
  const ext = createSPExtractor(item);

  let clickedItems: ISearchHistoryEntry['clickedItems'] = [];
  const rawClicked = ext.string('ClickedItems', '');
  if (rawClicked) {
    try {
      const parsed: unknown = JSON.parse(rawClicked);
      if (Array.isArray(parsed)) {
        clickedItems = parsed.map((c: Record<string, unknown>) => {
          const cExt = createSPExtractor(c);
          return {
            url: cExt.string('url', ''),
            title: cExt.string('title', ''),
            position: cExt.number('position', 0),
            timestamp: new Date((c.timestamp as string) || ''),
          };
        });
      }
    } catch {
      // Malformed JSON — ignore
    }
  }

  return {
    id: ext.number('Id', 0),
    queryHash: ext.string('QueryHash', ''),
    queryText: ext.string('Title', '') || ext.string('QueryText', ''),
    vertical: ext.string('Vertical', ''),
    scope: ext.string('Scope', ''),
    searchState: ext.string('SearchState', '{}'),
    resultCount: ext.number('ResultCount', 0),
    isZeroResult: ext.boolean('IsZeroResult', false),
    clickedItems,
    searchTimestamp: ext.date('SearchTimestamp') || ext.date('Created') || new Date(),
  };
}

function buildHistoryTitle(
  queryText: string,
  vertical: string,
  searchState: string
): string {
  const trimmedQuery = queryText.trim();
  if (trimmedQuery.length > 0) {
    return trimmedQuery.length > 255 ? trimmedQuery.substring(0, 255) : trimmedQuery;
  }

  let parsedState: IHistoryStateSummary | undefined;
  try {
    parsedState = JSON.parse(searchState) as IHistoryStateSummary;
  } catch {
    parsedState = undefined;
  }

  const filterLabels: string[] = [];
  const filters = parsedState?.activeFilters || [];
  for (let i = 0; i < filters.length && filterLabels.length < 2; i++) {
    const filter = filters[i];
    const label = filter.displayValue || filter.value || filter.filterName || '';
    if (label) {
      filterLabels.push((filter.filterName || 'Filter') + ': ' + label.replace(/^"|"$/g, ''));
    }
  }

  const verticalLabel = (parsedState?.currentVerticalKey || vertical || '').trim();

  let title = '';
  if (filterLabels.length > 0 && verticalLabel) {
    title = verticalLabel + ' • ' + filterLabels.join(' • ');
  } else if (filterLabels.length > 0) {
    title = filterLabels.join(' • ');
  } else if (verticalLabel) {
    title = 'Browse • ' + verticalLabel;
  } else {
    title = 'Browse all results';
  }

  return title.length > 255 ? title.substring(0, 255) : title;
}

/**
 * Maps raw SharePoint list items to ISearchCollection[].
 * Groups by CollectionName and includes real list item IDs.
 */
function mapToCollection(items: Array<Record<string, unknown>>): ISearchCollection[] {
  // Group by CollectionName
  const grouped: Map<string, Array<Record<string, unknown>>> = new Map();

  for (let i = 0; i < items.length; i++) {
    const ext = createSPExtractor(items[i]);
    const collectionName = ext.string('CollectionName', 'Untitled');
    const existing = grouped.get(collectionName);
    if (existing) {
      existing.push(items[i]);
    } else {
      grouped.set(collectionName, [items[i]]);
    }
  }

  const collections: ISearchCollection[] = [];

  grouped.forEach((groupItems, collectionName) => {
    const first = groupItems[0];
    const firstExt = createSPExtractor(first);

    // Map items WITH their real SharePoint list item IDs and per-item tags
    const collectionItems: ICollectionItem[] = [];
    const tagUnion: Set<string> = new Set();

    for (let i = 0; i < groupItems.length; i++) {
      const giExt = createSPExtractor(groupItems[i]);

      // Parse per-item tags
      let itemTags: string[] = [];
      const rawItemTags = giExt.string('Tags', '');
      if (rawItemTags) {
        try {
          const parsedTags: unknown = JSON.parse(rawItemTags);
          if (Array.isArray(parsedTags)) {
            itemTags = parsedTags as string[];
          }
        } catch {
          // Malformed JSON — ignore
        }
      }

      // Add to collection-level tag union
      for (let t = 0; t < itemTags.length; t++) {
        tagUnion.add(itemTags[t]);
      }

      collectionItems.push({
        id: giExt.number('Id', 0),
        url: giExt.string('ItemUrl', ''),
        title: giExt.string('ItemTitle', '') || giExt.string('Title', ''),
        metadata: (() => {
          try {
            const raw = giExt.string('ItemMetadata', '');
            return raw ? JSON.parse(raw) as Record<string, unknown> : {};
          } catch {
            return {};
          }
        })(),
        sortOrder: giExt.number('SortOrder', 0),
        tags: itemTags,
      });
    }

    // Sort by SortOrder
    collectionItems.sort((a, b) => a.sortOrder - b.sortOrder);

    collections.push({
      id: firstExt.number('Id', 0),
      collectionName,
      items: collectionItems,
      sharedWith: extractSharedWith(first),
      tags: Array.from(tagUnion),
      created: firstExt.date('Created') || new Date(),
      author: extractAuthor(first),
    });
  });

  return collections;
}

/**
 * SearchManagerService — CRUD operations for saved searches, collections,
 * and search history using SharePoint hidden lists.
 *
 * Uses SPContext.sp from spfx-toolkit for all SharePoint operations.
 * No SPFI constructor injection needed — SPContext is initialized once
 * in the web part onInit and available globally.
 *
 * CRITICAL: SearchHistory list WILL exceed 5,000 items.
 * All queries MUST use CAML with Author/AuthorId as the FIRST filter predicate.
 */
export class SearchManagerService {
  private _currentUserId: number = 0;
  private _initFailed: boolean = false;

  /**
   * Initialize the service by resolving the current user ID.
   * The user ID is critical for CAML queries that use AuthorId.
   *
   * If the current user cannot be resolved, the service enters a degraded
   * state where read operations return empty results and write operations
   * are silently skipped (prevents creating orphaned items).
   */
  public async initialize(): Promise<void> {
    // Reset failure state to allow retries after transient errors
    this._initFailed = false;
    this._currentUserId = 0;

    try {
      const user = await SPContext.sp.web.currentUser();
      this._currentUserId = user.Id || 0;
      if (this._currentUserId === 0) {
        this._initFailed = true;
        SPContext.logger.warn('SearchManagerService: Current user ID resolved to 0');
      } else {
        SPContext.logger.info('SearchManagerService: Initialized', { userId: this._currentUserId });
        this._maybeAutoCleanupHistory().catch(function noop(): void { /* non-critical */ });
      }
    } catch (error) {
      this._initFailed = true;
      SPContext.logger.warn('SearchManagerService: Failed to resolve current user', { error });
    }
  }

  /**
   * Returns true if the service is ready for write operations.
   * False when currentUser resolution failed — prevents creating
   * orphaned items that the user can never retrieve.
   */
  public get isReady(): boolean {
    return this._currentUserId > 0 && !this._initFailed;
  }

  // ─── Saved Searches ────────────────────────────────────────

  /**
   * Load all saved searches for the current user (owned + shared with me).
   */
  public async loadSavedSearches(): Promise<ISavedSearch[]> {
    if (!this.isReady) {
      console.warn('[SP Search] loadSavedSearches skipped — service not ready (userId=' + this._currentUserId + ')');
      return [];
    }

    try {
      // Select fields: start with core fields + Author expand.
      // SharedWith expand is attempted separately to prevent failures
      // (e.g., UserMulti field not indexed, misconfigured, or empty)
      // from blocking the entire query.
      const coreSelectFields = [
        'Id', 'Title', 'QueryText', 'SearchState', 'SearchUrl',
        'EntryType', 'Category', 'ResultCount', 'LastUsed', 'Created',
        'Author/Id', 'Author/Title', 'Author/EMail'
      ].join(',');

      const fullSelectFields = coreSelectFields + ',SharedWith/Id,SharedWith/Title,SharedWith/EMail';

      // Query 1: Items owned by current user
      // Try with SharedWith expand first; fall back without if it fails
      // CacheNever() bypasses any PnPjs caching behavior to ensure fresh data
      let ownedItems: Array<Record<string, unknown>>;
      try {
        ownedItems = await SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
          .items
          .using(CacheNever())
          .select(fullSelectFields)
          .expand('Author', 'SharedWith')
          .filter('Author/Id eq ' + this._currentUserId)
          .orderBy('LastUsed', false)
          .top(200)() as Array<Record<string, unknown>>;
      } catch {
        // SharedWith expand failed — retry without it
        console.warn('[SP Search] loadSavedSearches: SharedWith expand failed on owned query, retrying without');
        ownedItems = await SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
          .items
          .using(CacheNever())
          .select(coreSelectFields)
          .expand('Author')
          .filter('Author/Id eq ' + this._currentUserId)
          .orderBy('LastUsed', false)
          .top(200)() as Array<Record<string, unknown>>;
      }

      // Query 2: Items shared with current user (non-fatal)
      let sharedItems: Array<Record<string, unknown>> = [];
      try {
        sharedItems = await SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
          .items
          .using(CacheNever())
          .select(fullSelectFields)
          .expand('Author', 'SharedWith')
          .filter('SharedWith/Id eq ' + this._currentUserId)
          .orderBy('LastUsed', false)
          .top(100)() as Array<Record<string, unknown>>;
      } catch {
        // SharedWith field query may fail if not indexed
      }

      // Merge and deduplicate by Id
      const allItems = ownedItems;
      const seenIds = new Set<number>();
      for (let i = 0; i < allItems.length; i++) {
        seenIds.add(allItems[i].Id as number);
      }
      for (let i = 0; i < sharedItems.length; i++) {
        const itemId = sharedItems[i].Id as number;
        if (!seenIds.has(itemId)) {
          allItems.push(sharedItems[i]);
          seenIds.add(itemId);
        }
      }

      // Exclude state snapshots from saved search lists
      const filteredItems = allItems.filter(function (item): boolean {
        const entryType = (item.EntryType as string | undefined) || '';
        return entryType !== 'StateSnapshot';
      });

      // Sort by LastUsed descending
      filteredItems.sort((a, b) => {
        const dateA = new Date((a.LastUsed as string) || (a.Created as string) || '');
        const dateB = new Date((b.LastUsed as string) || (b.Created as string) || '');
        return dateB.getTime() - dateA.getTime();
      });

      console.log('[SP Search] loadSavedSearches: loaded ' + filteredItems.length + ' items');
      return filteredItems.map(mapToSavedSearch);
    } catch (error) {
      console.error('[SP Search] loadSavedSearches FAILED:', error);
      return [];
    }
  }

  /**
   * Save a new search to the SearchSavedQueries list.
   */
  public async saveSearch(
    title: string,
    queryText: string,
    searchState: string,
    searchUrl: string,
    category: string,
    resultCount: number
  ): Promise<ISavedSearch> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    SPContext.logger.info('SearchManagerService: Saving search', { title });

    const result = await SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
      .items
      .add({
        Title: title,
        QueryText: queryText,
        SearchState: searchState,
        SearchUrl: searchUrl,
        EntryType: 'SavedSearch',
        Category: category,
        ResultCount: resultCount,
        LastUsed: new Date().toISOString(),
      });

    const addedItem = (result as { data?: Record<string, unknown> }).data || result;
    return mapToSavedSearch(addedItem as Record<string, unknown>);
  }

  /**
   * Update an existing saved search.
   */
  public async updateSavedSearch(
    id: number,
    updates: {
      title?: string;
      searchState?: string;
      searchUrl?: string;
      category?: string;
      resultCount?: number;
    }
  ): Promise<void> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    const payload: Record<string, unknown> = {};
    if (updates.title !== undefined) {
      payload.Title = updates.title;
    }
    if (updates.searchState !== undefined) {
      payload.SearchState = updates.searchState;
    }
    if (updates.searchUrl !== undefined) {
      payload.SearchUrl = updates.searchUrl;
    }
    if (updates.category !== undefined) {
      payload.Category = updates.category;
    }
    if (updates.resultCount !== undefined) {
      payload.ResultCount = updates.resultCount;
    }
    payload.LastUsed = new Date().toISOString();

    await SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
      .items.getById(id).update(payload);

    SPContext.logger.info('SearchManagerService: Updated saved search', { id });
  }

  /**
   * Delete a saved search.
   * Uses recycle() (send to recycle bin) instead of delete() for more
   * reliable behavior in PnPjs v3 — avoids X-HTTP-Method header issues.
   */
  public async deleteSavedSearch(id: number): Promise<void> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    if (id <= 0) {
      throw new Error('Invalid item ID: ' + id);
    }

    console.log('[SP Search] deleteSavedSearch: recycling item id=' + id + ' from list=' + SAVED_QUERIES_LIST);
    await SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
      .items.getById(id).recycle();

    console.log('[SP Search] deleteSavedSearch: item id=' + id + ' recycled successfully');
  }

  // ─── Search History ────────────────────────────────────────

  /**
   * Load search history for the current user.
   *
   * CRITICAL: Uses CAML with AuthorId as the FIRST predicate to avoid
   * list view threshold issues on lists exceeding 5,000 items.
   */
  public async loadHistory(maxItems: number = 50): Promise<ISearchHistoryEntry[]> {
    if (!this.isReady) {
      console.warn('[SP Search] loadHistory skipped — service not ready (userId=' + this._currentUserId + ')');
      return [];
    }

    try {
      const camlQuery = `
        <View>
          <Query>
            <Where>
              <Eq>
                <FieldRef Name="Author" LookupId="TRUE" />
                <Value Type="Integer">${this._currentUserId}</Value>
              </Eq>
            </Where>
            <OrderBy>
              <FieldRef Name="SearchTimestamp" Ascending="FALSE" />
            </OrderBy>
          </Query>
          <RowLimit>${maxItems}</RowLimit>
          <ViewFields>
            <FieldRef Name="Id" />
            <FieldRef Name="Title" />
            <FieldRef Name="QueryHash" />
            <FieldRef Name="Vertical" />
            <FieldRef Name="Scope" />
            <FieldRef Name="SearchState" />
            <FieldRef Name="ResultCount" />
            <FieldRef Name="IsZeroResult" />
            <FieldRef Name="ClickedItems" />
            <FieldRef Name="SearchTimestamp" />
            <FieldRef Name="Created" />
          </ViewFields>
        </View>
      `;

      const items = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .getItemsByCAMLQuery({ ViewXml: camlQuery });

      console.log('[SP Search] loadHistory: loaded ' + (items as Array<unknown>).length + ' items');
      return (items as Array<Record<string, unknown>>).map(mapToHistoryEntry);
    } catch (error) {
      console.error('[SP Search] loadHistory FAILED:', error);
      return [];
    }
  }

  /**
   * Log a search to the history list (async, non-blocking).
   * Uses full state hash for deduplication.
   *
   * @returns The history entry ID (for click tracking)
   */
  public async logSearch(
    queryText: string,
    vertical: string,
    scope: string,
    searchState: string,
    resultCount: number,
    isZeroResult?: boolean
  ): Promise<number> {
    if (!this.isReady) {
      return 0;
    }
    try {
      const queryHash = await computeStateHash(searchState);
      const historyTitle = buildHistoryTitle(queryText, vertical, searchState);

      // Check for existing entry with same hash using CAML (Author-first)
      const camlQuery = `
        <View>
          <Query>
            <Where>
              <And>
                <Eq>
                  <FieldRef Name="Author" LookupId="TRUE" />
                  <Value Type="Integer">${this._currentUserId}</Value>
                </Eq>
                <Eq>
                  <FieldRef Name="QueryHash" />
                  <Value Type="Text">${this._escapeXmlValue(queryHash)}</Value>
                </Eq>
              </And>
            </Where>
          </Query>
          <RowLimit>1</RowLimit>
          <ViewFields>
            <FieldRef Name="Id" />
          </ViewFields>
        </View>
      `;

      const existing = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .getItemsByCAMLQuery({ ViewXml: camlQuery });

      if (existing && existing.length > 0) {
        const existingId = (existing[0] as Record<string, unknown>).Id as number;
        const updatePayload: Record<string, unknown> = {
          Title: historyTitle,
          QueryText: queryText,
          ResultCount: resultCount,
          IsZeroResult: isZeroResult === true,
          SearchTimestamp: new Date().toISOString(),
        };
        await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
          .items.getById(existingId).update(updatePayload);
        return existingId;
      } else {
        const addPayload: Record<string, unknown> = {
          Title: historyTitle,
          QueryText: queryText,
          QueryHash: queryHash,
          Vertical: vertical,
          Scope: scope,
          SearchState: searchState,
          ResultCount: resultCount,
          IsZeroResult: isZeroResult === true,
          SearchTimestamp: new Date().toISOString(),
        };
        const result = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
          .items.add(addPayload);
        const addedItem = (result as { data?: Record<string, unknown> }).data || result;
        return (addedItem as Record<string, unknown>).Id as number || 0;
      }
    } catch (error) {
      // Non-critical — log but don't throw history errors
      console.warn('SearchManagerService.logSearch failed:', error);
      return 0;
    }
  }

  private async _maybeAutoCleanupHistory(): Promise<void> {
    if (!this.isReady) {
      return;
    }

    const cleanupKey = 'sp-search-history-cleanup:' + SPContext.webAbsoluteUrl + ':' + String(this._currentUserId);
    let lastRun = 0;

    try {
      lastRun = parseInt(localStorage.getItem(cleanupKey) || '0', 10) || 0;
    } catch {
      lastRun = 0;
    }

    const now = Date.now();
    if (now - lastRun < HISTORY_CLEANUP_INTERVAL_MS) {
      return;
    }

    const deleted = await this.cleanupHistory(HISTORY_RETENTION_DAYS);

    try {
      localStorage.setItem(cleanupKey, String(now));
    } catch {
      // ignore
    }

    if (deleted > 0) {
      SPContext.logger.info('SearchManagerService: Auto-cleaned history', {
        deleted,
        ttlDays: HISTORY_RETENTION_DAYS
      });
    }
  }

  /**
   * Log a clicked item against a history entry.
   */
  public async logClickedItem(
    historyId: number,
    clickedUrl: string,
    clickedTitle: string,
    position: number
  ): Promise<void> {
    if (historyId <= 0) {
      return;
    }

    try {
      const item = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .items.getById(historyId)
        .select('ClickedItems')();

      const existing: Array<Record<string, unknown>> = [];
      const rawClicked = (item as Record<string, unknown>).ClickedItems as string;
      if (rawClicked) {
        try {
          const parsed: unknown = JSON.parse(rawClicked);
          if (Array.isArray(parsed)) {
            for (let i = 0; i < parsed.length; i++) {
              existing.push(parsed[i] as Record<string, unknown>);
            }
          }
        } catch {
          // Malformed JSON — start fresh
        }
      }

      existing.push({
        url: clickedUrl,
        title: clickedTitle,
        position,
        timestamp: new Date().toISOString(),
      });

      await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .items.getById(historyId).update({
          ClickedItems: JSON.stringify(existing),
        });
    } catch (error) {
      // Non-critical — ClickedItems field may not exist yet
      console.warn('SearchManagerService.logClickedItem failed:', error);
    }
  }

  /**
   * Clear all search history for the current user.
   * Uses BatchBuilder from spfx-toolkit for efficient batched deletes.
   */
  public async clearHistory(): Promise<void> {
    if (!this.isReady) {
      return;
    }
    try {
      let hasMore = true;
      while (hasMore) {
        const camlQuery = `
          <View>
            <Query>
              <Where>
                <Eq>
                  <FieldRef Name="Author" LookupId="TRUE" />
                  <Value Type="Integer">${this._currentUserId}</Value>
                </Eq>
              </Where>
            </Query>
            <RowLimit>100</RowLimit>
            <ViewFields>
              <FieldRef Name="Id" />
            </ViewFields>
          </View>
        `;

        const items = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
          .getItemsByCAMLQuery({ ViewXml: camlQuery });

        if (!items || items.length === 0) {
          hasMore = false;
          break;
        }

        // Use BatchBuilder from spfx-toolkit
        const batch = new BatchBuilder(SPContext.sp, { batchSize: 100 });
        const listOps = batch.list(HISTORY_LIST);

        for (let i = 0; i < items.length; i++) {
          const itemId = (items[i] as Record<string, unknown>).Id as number;
          listOps.delete(itemId);
        }

        await batch.execute();

        if (items.length < 100) {
          hasMore = false;
        }
      }

      SPContext.logger.info('SearchManagerService: Cleared history');
    } catch (error) {
      SPContext.logger.error('SearchManagerService: Failed to clear history', error);
    }
  }

  /**
   * Delete a single search history entry for the current user.
   * Verifies ownership via Author-first CAML before recycling the item.
   */
  public async deleteHistoryEntry(id: number): Promise<void> {
    if (!this.isReady || id <= 0) {
      return;
    }

    try {
      const camlQuery = `
        <View>
          <Query>
            <Where>
              <And>
                <Eq>
                  <FieldRef Name="Author" LookupId="TRUE" />
                  <Value Type="Integer">${this._currentUserId}</Value>
                </Eq>
                <Eq>
                  <FieldRef Name="Id" />
                  <Value Type="Counter">${id}</Value>
                </Eq>
              </And>
            </Where>
          </Query>
          <RowLimit>1</RowLimit>
          <ViewFields>
            <FieldRef Name="Id" />
          </ViewFields>
        </View>
      `;

      const items = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .getItemsByCAMLQuery({ ViewXml: camlQuery });

      if (!items || items.length === 0) {
        return;
      }

      await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .items.getById(id)
        .recycle();
    } catch (error) {
      SPContext.logger.error('SearchManagerService: Failed to delete history entry', error);
    }
  }

  // ─── Collections ───────────────────────────────────────────

  /**
   * Load all collections for the current user (owned + shared with me).
   */
  public async loadCollections(): Promise<ISearchCollection[]> {
    if (!this.isReady) {
      console.warn('[SP Search] loadCollections skipped — service not ready (userId=' + this._currentUserId + ')');
      return [];
    }

    try {
      const coreSelectFields = [
        'Id', 'Title', 'ItemUrl', 'ItemTitle', 'ItemMetadata',
        'CollectionName', 'Tags', 'SortOrder', 'Created',
        'Author/Id', 'Author/Title', 'Author/EMail'
      ].join(',');

      const fullSelectFields = coreSelectFields + ',SharedWith/Id,SharedWith/Title,SharedWith/EMail';

      // Query 1: Items owned by current user
      // Try with SharedWith expand first; fall back without if it fails
      let ownedItems: Array<Record<string, unknown>>;
      try {
        ownedItems = await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
          .items
          .using(CacheNever())
          .select(fullSelectFields)
          .expand('Author', 'SharedWith')
          .filter('Author/Id eq ' + this._currentUserId)
          .orderBy('CollectionName', true)
          .top(500)() as Array<Record<string, unknown>>;
      } catch {
        console.warn('[SP Search] loadCollections: SharedWith expand failed on owned query, retrying without');
        ownedItems = await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
          .items
          .using(CacheNever())
          .select(coreSelectFields)
          .expand('Author')
          .filter('Author/Id eq ' + this._currentUserId)
          .orderBy('CollectionName', true)
          .top(500)() as Array<Record<string, unknown>>;
      }

      // Query 2: Items shared with current user (non-fatal)
      let sharedItems: Array<Record<string, unknown>> = [];
      try {
        sharedItems = await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
          .items
          .using(CacheNever())
          .select(fullSelectFields)
          .expand('Author', 'SharedWith')
          .filter('SharedWith/Id eq ' + this._currentUserId)
          .orderBy('CollectionName', true)
          .top(200)() as Array<Record<string, unknown>>;
      } catch {
        // SharedWith field query may fail if not indexed
      }

      // Merge and deduplicate by Id
      const allItems = ownedItems;
      const seenIds = new Set<number>();
      for (let i = 0; i < allItems.length; i++) {
        seenIds.add(allItems[i].Id as number);
      }
      for (let i = 0; i < sharedItems.length; i++) {
        const itemId = sharedItems[i].Id as number;
        if (!seenIds.has(itemId)) {
          allItems.push(sharedItems[i]);
          seenIds.add(itemId);
        }
      }

      console.log('[SP Search] loadCollections: loaded ' + allItems.length + ' items');
      return mapToCollection(allItems);
    } catch (error) {
      console.error('[SP Search] loadCollections FAILED:', error);
      return [];
    }
  }

  /**
   * Create a new collection.
   */
  public async createCollection(name: string): Promise<void> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .items.add({
        Title: name,
        CollectionName: name,
        ItemUrl: '',
        ItemTitle: '',
        SortOrder: 0,
        Tags: '[]',
      });

    SPContext.logger.info('SearchManagerService: Created collection', { name });
  }

  /**
   * Pin a result to a collection.
   */
  public async pinToCollection(
    collectionName: string,
    itemUrl: string,
    itemTitle: string,
    metadata: Record<string, unknown>,
    tags?: string[]
  ): Promise<void> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    const camlQuery = `
      <View>
        <Query>
          <Where>
            <And>
              <Eq>
                <FieldRef Name="Author" LookupId="TRUE" />
                <Value Type="Integer">${this._currentUserId}</Value>
              </Eq>
              <Eq>
                <FieldRef Name="CollectionName" />
                <Value Type="Text">${this._escapeXmlValue(collectionName)}</Value>
              </Eq>
            </And>
          </Where>
          <OrderBy>
            <FieldRef Name="SortOrder" Ascending="FALSE" />
          </OrderBy>
        </Query>
        <RowLimit>1</RowLimit>
        <ViewFields>
          <FieldRef Name="SortOrder" />
        </ViewFields>
      </View>
    `;

    const existing = await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .getItemsByCAMLQuery({ ViewXml: camlQuery });

    const maxOrder = existing.length > 0
      ? ((existing[0] as Record<string, unknown>).SortOrder as number || 0) + 1
      : 0;

    await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .items.add({
        Title: itemTitle,
        CollectionName: collectionName,
        ItemUrl: itemUrl,
        ItemTitle: itemTitle,
        ItemMetadata: JSON.stringify(metadata),
        SortOrder: maxOrder,
        Tags: JSON.stringify(tags || []),
      });

    SPContext.logger.info('SearchManagerService: Pinned item to collection', { collectionName, itemTitle });
  }

  /**
   * Update the tags on a specific collection item.
   * Tags are stored as a JSON array in the Tags column.
   */
  public async updateItemTags(itemId: number, tags: string[]): Promise<void> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    // Sanitize: trim, remove empties, deduplicate, limit length
    const sanitized: string[] = [];
    const seen: Set<string> = new Set();
    for (let i = 0; i < tags.length; i++) {
      const trimmed = tags[i].trim();
      if (trimmed.length > 0 && trimmed.length <= 50 && !seen.has(trimmed)) {
        sanitized.push(trimmed);
        seen.add(trimmed);
      }
      if (sanitized.length >= 20) {
        break;
      }
    }

    await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .items.getById(itemId).update({
        Tags: JSON.stringify(sanitized),
      });

    SPContext.logger.info('SearchManagerService: Updated item tags', { itemId, tagCount: sanitized.length });
  }

  /**
   * Remove a pinned item from a collection.
   */
  public async unpinFromCollection(itemId: number): Promise<void> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .items.getById(itemId).recycle();
  }

  /**
   * Delete an entire collection using BatchBuilder.
   */
  public async deleteCollection(collectionName: string): Promise<void> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    const camlQuery = `
      <View>
        <Query>
          <Where>
            <And>
              <Eq>
                <FieldRef Name="Author" LookupId="TRUE" />
                <Value Type="Integer">${this._currentUserId}</Value>
              </Eq>
              <Eq>
                <FieldRef Name="CollectionName" />
                <Value Type="Text">${this._escapeXmlValue(collectionName)}</Value>
              </Eq>
            </And>
          </Where>
        </Query>
        <RowLimit>500</RowLimit>
        <ViewFields>
          <FieldRef Name="Id" />
        </ViewFields>
      </View>
    `;

    const items = await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .getItemsByCAMLQuery({ ViewXml: camlQuery });

    if (items.length === 0) {
      return;
    }

    const batch = new BatchBuilder(SPContext.sp, { batchSize: 100 });
    const listOps = batch.list(COLLECTIONS_LIST);

    for (let i = 0; i < items.length; i++) {
      const itemId = (items[i] as Record<string, unknown>).Id as number;
      listOps.delete(itemId);
    }

    await batch.execute();
    SPContext.logger.info('SearchManagerService: Deleted collection', { collectionName, itemCount: items.length });
  }

  /**
   * Rename a collection using BatchBuilder.
   */
  public async renameCollection(oldName: string, newName: string): Promise<void> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }
    const camlQuery = `
      <View>
        <Query>
          <Where>
            <And>
              <Eq>
                <FieldRef Name="Author" LookupId="TRUE" />
                <Value Type="Integer">${this._currentUserId}</Value>
              </Eq>
              <Eq>
                <FieldRef Name="CollectionName" />
                <Value Type="Text">${this._escapeXmlValue(oldName)}</Value>
              </Eq>
            </And>
          </Where>
        </Query>
        <RowLimit>500</RowLimit>
        <ViewFields>
          <FieldRef Name="Id" />
        </ViewFields>
      </View>
    `;

    const items = await SPContext.sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .getItemsByCAMLQuery({ ViewXml: camlQuery });

    const batch = new BatchBuilder(SPContext.sp, { batchSize: 100 });
    const listOps = batch.list(COLLECTIONS_LIST);

    for (let i = 0; i < items.length; i++) {
      const itemId = (items[i] as Record<string, unknown>).Id as number;
      listOps.update(itemId, { CollectionName: newName });
    }

    await batch.execute();
    SPContext.logger.info('SearchManagerService: Renamed collection', { oldName, newName });
  }

  // ─── Sharing ─────────────────────────────────────────────

  /**
   * Share a saved search with specific users.
   * Updates SharedWith field and sets item-level read permissions.
   */
  public async shareToUsers(savedSearchId: number, userEmails: string[]): Promise<void> {
    if (userEmails.length === 0) {
      return;
    }
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }

    SPContext.logger.info('SearchManagerService: Sharing search', { savedSearchId, userCount: userEmails.length });

    const list = SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST);
    const item = list.items.getById(savedSearchId);

    // Resolve user IDs
    const userIds: number[] = [];
    for (let i = 0; i < userEmails.length; i++) {
      try {
        const user = await SPContext.sp.web.ensureUser(userEmails[i]);
        if (user.data && user.data.Id) {
          userIds.push(user.data.Id);
        }
      } catch {
        SPContext.logger.warn('SearchManagerService: Could not resolve user', { email: userEmails[i] });
      }
    }

    if (userIds.length === 0) {
      return;
    }

    // Update SharedWith field
    await item.update({
      SharedWithId: { results: userIds },
      EntryType: 'SharedSearch',
    });

    // Break role inheritance and grant read access
    try {
      await item.breakRoleInheritance(true, false);

      const READ_ROLE_DEF_ID = 1073741826;
      for (let i = 0; i < userIds.length; i++) {
        await item.roleAssignments.add(userIds[i], READ_ROLE_DEF_ID);
      }
    } catch {
      // Permission operations may fail; SharedWith update still succeeded
      SPContext.logger.warn('SearchManagerService: Could not set item-level permissions', { savedSearchId });
    }
  }

  // ─── History Cleanup ────────────────────────────────────────

  /**
   * Delete history entries older than the specified number of days.
   * Uses BatchBuilder for efficient batched deletes.
   */
  public async cleanupHistory(ttlDays: number): Promise<number> {
    if (!this.isReady) {
      return 0;
    }
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - ttlDays);
    const cutoffIso = cutoff.toISOString();

    let totalDeleted = 0;
    let hasMore = true;

    while (hasMore) {
      const camlQuery = `
        <View>
          <Query>
            <Where>
              <And>
                <Eq>
                  <FieldRef Name="Author" LookupId="TRUE" />
                  <Value Type="Integer">${this._currentUserId}</Value>
                </Eq>
                <Lt>
                  <FieldRef Name="SearchTimestamp" />
                  <Value Type="DateTime" IncludeTimeValue="TRUE">${cutoffIso}</Value>
                </Lt>
              </And>
            </Where>
          </Query>
          <RowLimit>100</RowLimit>
          <ViewFields>
            <FieldRef Name="Id" />
          </ViewFields>
        </View>
      `;

      const items = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .getItemsByCAMLQuery({ ViewXml: camlQuery });

      if (!items || items.length === 0) {
        hasMore = false;
        break;
      }

      const batch = new BatchBuilder(SPContext.sp, { batchSize: 100 });
      const listOps = batch.list(HISTORY_LIST);

      for (let i = 0; i < items.length; i++) {
        const itemId = (items[i] as Record<string, unknown>).Id as number;
        listOps.delete(itemId);
      }

      await batch.execute();
      totalDeleted += items.length;

      if (items.length < 100) {
        hasMore = false;
      }
    }

    SPContext.logger.info('SearchManagerService: History cleanup complete', { ttlDays, totalDeleted });
    return totalDeleted;
  }

  // ─── Zero-Result Query Health ─────────────────────────────────

  /**
   * Load recent zero-result queries across ALL users (admin health view).
   *
   * Uses SearchTimestamp as the first CAML predicate (indexed) to avoid the
   * 5,000-item list-view threshold. IsZeroResult = 1 is the second predicate.
   * No Author filter — this is intentionally an admin-level cross-user query.
   *
   * @param daysBack - Window to scan (default 90 days)
   * @param maxItems - Row cap (default 200; client aggregates before display)
   */
  public async loadZeroResultQueries(
    daysBack: number = 90,
    maxItems: number = 200
  ): Promise<ISearchHistoryEntry[]> {
    if (!this.isReady) {
      return [];
    }
    try {
      const cutoff = new Date();
      cutoff.setDate(cutoff.getDate() - daysBack);
      const cutoffIso = cutoff.toISOString();

      const camlQuery = `
        <View>
          <Query>
            <Where>
              <And>
                <Geq>
                  <FieldRef Name="SearchTimestamp" />
                  <Value Type="DateTime" IncludeTimeValue="TRUE">${cutoffIso}</Value>
                </Geq>
                <Eq>
                  <FieldRef Name="IsZeroResult" />
                  <Value Type="Boolean">1</Value>
                </Eq>
              </And>
            </Where>
            <OrderBy>
              <FieldRef Name="SearchTimestamp" Ascending="FALSE" />
            </OrderBy>
          </Query>
          <RowLimit>${maxItems}</RowLimit>
          <ViewFields>
            <FieldRef Name="Id" />
            <FieldRef Name="Title" />
            <FieldRef Name="QueryHash" />
            <FieldRef Name="Vertical" />
            <FieldRef Name="Scope" />
            <FieldRef Name="SearchState" />
            <FieldRef Name="ResultCount" />
            <FieldRef Name="IsZeroResult" />
            <FieldRef Name="SearchTimestamp" />
            <FieldRef Name="Created" />
          </ViewFields>
        </View>
      `;

      const items = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .getItemsByCAMLQuery({ ViewXml: camlQuery });

      return (items as Array<Record<string, unknown>>).map(mapToHistoryEntry);
    } catch (error) {
      console.warn('SearchManagerService.loadZeroResultQueries failed:', error);
      return [];
    }
  }

  /**
   * Load all search history events within a date window across ALL users.
   * Used by the Insights panel to compute aggregate metrics client-side.
   *
   * Uses SearchTimestamp as the first (indexed) predicate — safe above 5k items.
   * No Author filter — intentional admin-level cross-user query.
   *
   * @param daysBack - Window to scan (default 30 days)
   * @param maxItems - Row cap before client aggregation (default 500)
   */
  public async loadAllHistoryForInsights(
    daysBack: number = 30,
    maxItems: number = 500
  ): Promise<ISearchHistoryEntry[]> {
    if (!this.isReady) {
      return [];
    }
    try {
      const cutoff = new Date();
      cutoff.setDate(cutoff.getDate() - daysBack);
      const cutoffIso = cutoff.toISOString();

      const camlQuery = `
        <View>
          <Query>
            <Where>
              <Geq>
                <FieldRef Name="SearchTimestamp" />
                <Value Type="DateTime" IncludeTimeValue="TRUE">${cutoffIso}</Value>
              </Geq>
            </Where>
            <OrderBy>
              <FieldRef Name="SearchTimestamp" Ascending="FALSE" />
            </OrderBy>
          </Query>
          <RowLimit>${maxItems}</RowLimit>
          <ViewFields>
            <FieldRef Name="Id" />
            <FieldRef Name="Title" />
            <FieldRef Name="Vertical" />
            <FieldRef Name="ResultCount" />
            <FieldRef Name="IsZeroResult" />
            <FieldRef Name="ClickedItems" />
            <FieldRef Name="SearchTimestamp" />
          </ViewFields>
        </View>
      `;

      const items = await SPContext.sp.web.lists.getByTitle(HISTORY_LIST)
        .getItemsByCAMLQuery({ ViewXml: camlQuery });

      return (items as Array<Record<string, unknown>>).map(mapToHistoryEntry);
    } catch (error) {
      console.warn('SearchManagerService.loadAllHistoryForInsights failed:', error);
      return [];
    }
  }

  // ─── State Snapshots (StateId Fallback) ─────────────────────

  /**
   * Save a full state snapshot to SearchSavedQueries list.
   * Stored as EntryType = StateSnapshot, SearchState = JSON.
   * @returns The list item ID to use as ?sid= parameter
   */
  public async saveStateSnapshot(stateJson: string): Promise<number> {
    if (!this.isReady) {
      throw new Error('SearchManagerService is not ready — current user could not be resolved');
    }

    const expiresAt = new Date();
    expiresAt.setDate(expiresAt.getDate() + 30);

    const result = await SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
      .items.add({
        Title: 'StateSnapshot-' + new Date().getTime(),
        EntryType: 'StateSnapshot',
        SearchState: stateJson,
        SearchUrl: '',
        QueryText: '',
        Category: '',
        ResultCount: 0,
        LastUsed: new Date().toISOString(),
        ExpiresAt: expiresAt.toISOString(),
      });

    const addedItem = (result as { data?: Record<string, unknown> }).data || result;
    return (addedItem as Record<string, unknown>).Id as number || 0;
  }

  /**
   * Load a state snapshot by item ID.
   */
  public async loadStateSnapshot(stateId: number): Promise<string> {
    try {
      const item = await SPContext.sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
        .items.getById(stateId)
        .select('SearchState', 'EntryType', 'ExpiresAt')();

      const entryType = (item as Record<string, unknown>).EntryType as string | undefined;
      if (entryType !== 'StateSnapshot') {
        return '';
      }

      const expiresAtRaw = (item as Record<string, unknown>).ExpiresAt as string | undefined;
      if (expiresAtRaw) {
        const expiresAt = new Date(expiresAtRaw);
        if (!isNaN(expiresAt.getTime()) && new Date() > expiresAt) {
          return '';
        }
      }

      return ((item as Record<string, unknown>).SearchState as string) || '';
    } catch {
      return '';
    }
  }

  // ─── Helpers ───────────────────────────────────────────────

  /**
   * Escape a string for CAML XML values.
   */
  private _escapeXmlValue(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
