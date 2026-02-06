import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/batching';
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
 * Maps a raw SharePoint list item to ISavedSearch.
 */
function mapToSavedSearch(item: Record<string, unknown>): ISavedSearch {
  const sharedWith: IPersonaInfo[] = [];
  const rawShared = item.SharedWith as Array<Record<string, unknown>> | undefined;
  if (rawShared && Array.isArray(rawShared)) {
    for (let i = 0; i < rawShared.length; i++) {
      sharedWith.push({
        displayText: (rawShared[i].Title as string) || '',
        email: (rawShared[i].EMail as string) || '',
        imageUrl: undefined,
      });
    }
  }

  const authorRaw = item.Author as Record<string, unknown> | undefined;

  return {
    id: item.Id as number || 0,
    title: (item.Title as string) || '',
    queryText: (item.QueryText as string) || '',
    searchState: (item.SearchState as string) || '{}',
    searchUrl: (item.SearchUrl as string) || '',
    entryType: ((item.EntryType as string) || 'SavedSearch') as 'SavedSearch' | 'SharedSearch',
    category: (item.Category as string) || '',
    sharedWith,
    resultCount: (item.ResultCount as number) || 0,
    lastUsed: new Date((item.LastUsed as string) || (item.Created as string) || ''),
    created: new Date((item.Created as string) || ''),
    author: {
      displayText: authorRaw ? (authorRaw.Title as string || '') : '',
      email: authorRaw ? (authorRaw.EMail as string || '') : '',
      imageUrl: undefined,
    },
  };
}

/**
 * Maps a raw SharePoint list item to ISearchHistoryEntry.
 */
function mapToHistoryEntry(item: Record<string, unknown>): ISearchHistoryEntry {
  let clickedItems: ISearchHistoryEntry['clickedItems'] = [];
  const rawClicked = item.ClickedItems as string | undefined;
  if (rawClicked) {
    try {
      const parsed: unknown = JSON.parse(rawClicked);
      if (Array.isArray(parsed)) {
        clickedItems = parsed.map((c: Record<string, unknown>) => ({
          url: (c.url as string) || '',
          title: (c.title as string) || '',
          position: (c.position as number) || 0,
          timestamp: new Date((c.timestamp as string) || ''),
        }));
      }
    } catch {
      // Malformed JSON — ignore
    }
  }

  return {
    id: (item.Id as number) || 0,
    queryHash: (item.QueryHash as string) || '',
    queryText: (item.Title as string) || '',
    vertical: (item.Vertical as string) || '',
    scope: (item.Scope as string) || '',
    searchState: (item.SearchState as string) || '{}',
    resultCount: (item.ResultCount as number) || 0,
    clickedItems,
    searchTimestamp: new Date((item.SearchTimestamp as string) || (item.Created as string) || ''),
  };
}

/**
 * Maps raw SharePoint list items to ISearchCollection[].
 * Groups by CollectionName and includes real list item IDs.
 */
function mapToCollection(items: Array<Record<string, unknown>>): ISearchCollection[] {
  // Group by CollectionName
  const grouped: Map<string, Array<Record<string, unknown>>> = new Map();

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const collectionName = (item.CollectionName as string) || 'Untitled';
    const existing = grouped.get(collectionName);
    if (existing) {
      existing.push(item);
    } else {
      grouped.set(collectionName, [item]);
    }
  }

  const collections: ISearchCollection[] = [];

  grouped.forEach((groupItems, collectionName) => {
    // Use first item's metadata for collection-level properties
    const first = groupItems[0];

    const sharedWith: IPersonaInfo[] = [];
    const rawShared = first.SharedWith as Array<Record<string, unknown>> | undefined;
    if (rawShared && Array.isArray(rawShared)) {
      for (let i = 0; i < rawShared.length; i++) {
        sharedWith.push({
          displayText: (rawShared[i].Title as string) || '',
          email: (rawShared[i].EMail as string) || '',
          imageUrl: undefined,
        });
      }
    }

    let tags: string[] = [];
    const rawTags = first.Tags as string | undefined;
    if (rawTags) {
      try {
        const parsed: unknown = JSON.parse(rawTags);
        if (Array.isArray(parsed)) {
          tags = parsed as string[];
        }
      } catch {
        // Malformed JSON — ignore
      }
    }

    const authorRaw = first.Author as Record<string, unknown> | undefined;

    // Map items WITH their real SharePoint list item IDs
    const collectionItems: ICollectionItem[] = groupItems.map((gi) => ({
      id: (gi.Id as number) || 0, // Real list item ID for unpin
      url: (gi.ItemUrl as string) || '',
      title: (gi.ItemTitle as string) || (gi.Title as string) || '',
      metadata: (() => {
        try {
          const raw = gi.ItemMetadata as string;
          return raw ? JSON.parse(raw) as Record<string, unknown> : {};
        } catch {
          return {};
        }
      })(),
      sortOrder: (gi.SortOrder as number) || 0,
    }));

    // Sort by SortOrder
    collectionItems.sort((a, b) => a.sortOrder - b.sortOrder);

    collections.push({
      id: (first.Id as number) || 0,
      collectionName,
      items: collectionItems,
      sharedWith,
      tags,
      created: new Date((first.Created as string) || ''),
      author: {
        displayText: authorRaw ? (authorRaw.Title as string || '') : '',
        email: authorRaw ? (authorRaw.EMail as string || '') : '',
        imageUrl: undefined,
      },
    });
  });

  return collections;
}

/**
 * SearchManagerService — CRUD operations for saved searches, collections,
 * and search history using SharePoint hidden lists.
 *
 * CRITICAL: SearchHistory list WILL exceed 5,000 items.
 * All queries MUST use CAML with Author/AuthorId as the FIRST filter predicate.
 */
export class SearchManagerService {
  private readonly _sp: SPFI;
  private _currentUserId: number = 0;

  public constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Initialize the service by resolving the current user ID and email.
   * The user ID is critical for CAML queries that use AuthorId.
   */
  public async initialize(): Promise<void> {
    try {
      const user = await this._sp.web.currentUser();
      this._currentUserId = user.Id || 0;
    } catch {
      // Fallback — will use 0 for user ID
    }
  }

  // ─── Saved Searches ────────────────────────────────────────

  /**
   * Load all saved searches for the current user (owned + shared with me).
   * Queries owned items and items shared with the user, then merges.
   */
  public async loadSavedSearches(): Promise<ISavedSearch[]> {
    try {
      const selectFields = [
        'Id', 'Title', 'QueryText', 'SearchState', 'SearchUrl',
        'EntryType', 'Category', 'ResultCount', 'LastUsed', 'Created',
        'Author/Id', 'Author/Title', 'Author/EMail',
        'SharedWith/Id', 'SharedWith/Title', 'SharedWith/EMail'
      ].join(',');

      // Query 1: Items owned by current user
      const ownedItems = await this._sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
        .items
        .select(selectFields)
        .expand('Author', 'SharedWith')
        .filter(`Author/Id eq ${this._currentUserId}`)
        .orderBy('LastUsed', false)
        .top(200)();

      // Query 2: Items shared with current user (where user is in SharedWith field)
      // SharedWith is a multi-value Person field - use the LookupId sub-property
      let sharedItems: Array<Record<string, unknown>> = [];
      try {
        sharedItems = await this._sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
          .items
          .select(selectFields)
          .expand('Author', 'SharedWith')
          .filter(`SharedWith/Id eq ${this._currentUserId}`)
          .orderBy('LastUsed', false)
          .top(100)() as Array<Record<string, unknown>>;
      } catch {
        // SharedWith field query may fail if not indexed; fall back to owned only
      }

      // Merge and deduplicate by Id
      const allItems = ownedItems as Array<Record<string, unknown>>;
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

      // Sort by LastUsed descending
      allItems.sort((a, b) => {
        const dateA = new Date((a.LastUsed as string) || (a.Created as string) || '');
        const dateB = new Date((b.LastUsed as string) || (b.Created as string) || '');
        return dateB.getTime() - dateA.getTime();
      });

      return allItems.map(mapToSavedSearch);
    } catch {
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
    const result = await this._sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
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

    // PnPjs v3 returns { data, item } - extract the data
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

    await this._sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
      .items.getById(id).update(payload);
  }

  /**
   * Delete a saved search.
   */
  public async deleteSavedSearch(id: number): Promise<void> {
    await this._sp.web.lists.getByTitle(SAVED_QUERIES_LIST)
      .items.getById(id).delete();
  }

  // ─── Search History ────────────────────────────────────────

  /**
   * Load search history for the current user.
   *
   * CRITICAL: Uses CAML with AuthorId as the FIRST predicate to avoid
   * list view threshold issues on lists exceeding 5,000 items.
   */
  public async loadHistory(maxItems: number = 50): Promise<ISearchHistoryEntry[]> {
    try {
      // Use CAML to ensure Author is the FIRST filter predicate
      // AuthorId is an indexed column which enables efficient filtering
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
            <FieldRef Name="ClickedItems" />
            <FieldRef Name="SearchTimestamp" />
            <FieldRef Name="Created" />
          </ViewFields>
        </View>
      `;

      const items = await this._sp.web.lists.getByTitle(HISTORY_LIST)
        .getItemsByCAMLQuery({ ViewXml: camlQuery });

      return (items as Array<Record<string, unknown>>).map(mapToHistoryEntry);
    } catch {
      return [];
    }
  }

  /**
   * Log a search to the history list (async, non-blocking).
   * Uses full state hash for deduplication (includes filters, sort, etc.).
   *
   * @returns The history entry ID (for click tracking)
   */
  public async logSearch(
    queryText: string,
    vertical: string,
    scope: string,
    searchState: string,
    resultCount: number
  ): Promise<number> {
    try {
      const queryHash = await computeStateHash(searchState);

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

      const existing = await this._sp.web.lists.getByTitle(HISTORY_LIST)
        .getItemsByCAMLQuery({ ViewXml: camlQuery });

      if (existing && existing.length > 0) {
        // Update existing entry's timestamp and result count
        const existingId = (existing[0] as Record<string, unknown>).Id as number;
        await this._sp.web.lists.getByTitle(HISTORY_LIST)
          .items.getById(existingId).update({
            ResultCount: resultCount,
            SearchTimestamp: new Date().toISOString(),
          });
        return existingId;
      } else {
        // Create new history entry
        const result = await this._sp.web.lists.getByTitle(HISTORY_LIST)
          .items.add({
            Title: queryText.length > 255 ? queryText.substring(0, 255) : queryText,
            QueryHash: queryHash,
            Vertical: vertical,
            Scope: scope,
            SearchState: searchState,
            ResultCount: resultCount,
            ClickedItems: '[]',
            SearchTimestamp: new Date().toISOString(),
          });
        const addedItem = (result as { data?: Record<string, unknown> }).data || result;
        return (addedItem as Record<string, unknown>).Id as number || 0;
      }
    } catch {
      // Non-critical — swallow history logging errors
      return 0;
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
      // Load existing clicked items
      const item = await this._sp.web.lists.getByTitle(HISTORY_LIST)
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

      await this._sp.web.lists.getByTitle(HISTORY_LIST)
        .items.getById(historyId).update({
          ClickedItems: JSON.stringify(existing),
        });
    } catch {
      // Non-critical — swallow errors
    }
  }

  /**
   * Clear all search history for the current user.
   * Uses batched delete for efficiency and handles large lists properly.
   */
  public async clearHistory(): Promise<void> {
    try {
      // Use CAML to get all history items for current user (Author-first)
      // Process in batches to handle large lists
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

        const items = await this._sp.web.lists.getByTitle(HISTORY_LIST)
          .getItemsByCAMLQuery({ ViewXml: camlQuery });

        if (!items || items.length === 0) {
          hasMore = false;
          break;
        }

        // Use batched delete for efficiency
        const [batchedSP, execute] = this._sp.batched();
        const list = batchedSP.web.lists.getByTitle(HISTORY_LIST);

        for (let i = 0; i < items.length; i++) {
          const itemId = (items[i] as Record<string, unknown>).Id as number;
          list.items.getById(itemId).delete();
        }

        await execute();

        // If we got less than 100, we're done
        if (items.length < 100) {
          hasMore = false;
        }
      }
    } catch {
      // Swallow errors
    }
  }

  // ─── Collections ───────────────────────────────────────────

  /**
   * Load all collections for the current user (owned + shared with me).
   * Queries owned items and items shared with the user, then merges.
   */
  public async loadCollections(): Promise<ISearchCollection[]> {
    try {
      const selectFields = [
        'Id', 'Title', 'ItemUrl', 'ItemTitle', 'ItemMetadata',
        'CollectionName', 'Tags', 'SortOrder', 'Created',
        'Author/Id', 'Author/Title', 'Author/EMail',
        'SharedWith/Id', 'SharedWith/Title', 'SharedWith/EMail'
      ].join(',');

      // Query 1: Items owned by current user
      const ownedItems = await this._sp.web.lists.getByTitle(COLLECTIONS_LIST)
        .items
        .select(selectFields)
        .expand('Author', 'SharedWith')
        .filter(`Author/Id eq ${this._currentUserId}`)
        .orderBy('CollectionName', true)
        .top(500)();

      // Query 2: Items shared with current user
      let sharedItems: Array<Record<string, unknown>> = [];
      try {
        sharedItems = await this._sp.web.lists.getByTitle(COLLECTIONS_LIST)
          .items
          .select(selectFields)
          .expand('Author', 'SharedWith')
          .filter(`SharedWith/Id eq ${this._currentUserId}`)
          .orderBy('CollectionName', true)
          .top(200)() as Array<Record<string, unknown>>;
      } catch {
        // SharedWith field query may fail if not indexed; fall back to owned only
      }

      // Merge and deduplicate by Id
      const allItems = ownedItems as Array<Record<string, unknown>>;
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

      return mapToCollection(allItems);
    } catch {
      return [];
    }
  }

  /**
   * Create a new collection.
   */
  public async createCollection(name: string): Promise<void> {
    // Create a placeholder item to establish the collection
    await this._sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .items.add({
        Title: name,
        CollectionName: name,
        ItemUrl: '',
        ItemTitle: '',
        SortOrder: 0,
        Tags: '[]',
      });
  }

  /**
   * Pin a result to a collection.
   */
  public async pinToCollection(
    collectionName: string,
    itemUrl: string,
    itemTitle: string,
    metadata: Record<string, unknown>
  ): Promise<void> {
    // Use CAML to get current max sort order for this collection (Author-first)
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

    const existing = await this._sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .getItemsByCAMLQuery({ ViewXml: camlQuery });

    const maxOrder = existing.length > 0
      ? ((existing[0] as Record<string, unknown>).SortOrder as number || 0) + 1
      : 0;

    await this._sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .items.add({
        Title: itemTitle,
        CollectionName: collectionName,
        ItemUrl: itemUrl,
        ItemTitle: itemTitle,
        ItemMetadata: JSON.stringify(metadata),
        SortOrder: maxOrder,
        Tags: '[]',
      });
  }

  /**
   * Remove a pinned item from a collection.
   * Uses the real SharePoint list item ID.
   */
  public async unpinFromCollection(itemId: number): Promise<void> {
    await this._sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .items.getById(itemId).delete();
  }

  /**
   * Delete an entire collection (all its items).
   * Uses batched delete for efficiency.
   */
  public async deleteCollection(collectionName: string): Promise<void> {
    // Use CAML with Author-first predicate
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

    const items = await this._sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .getItemsByCAMLQuery({ ViewXml: camlQuery });

    if (items.length === 0) {
      return;
    }

    // Use batched delete
    const [batchedSP, execute] = this._sp.batched();
    const list = batchedSP.web.lists.getByTitle(COLLECTIONS_LIST);

    for (let i = 0; i < items.length; i++) {
      const itemId = (items[i] as Record<string, unknown>).Id as number;
      list.items.getById(itemId).delete();
    }

    await execute();
  }

  /**
   * Rename a collection.
   */
  public async renameCollection(oldName: string, newName: string): Promise<void> {
    // Use CAML with Author-first predicate
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

    const items = await this._sp.web.lists.getByTitle(COLLECTIONS_LIST)
      .getItemsByCAMLQuery({ ViewXml: camlQuery });

    // Use batched update
    const [batchedSP, execute] = this._sp.batched();
    const list = batchedSP.web.lists.getByTitle(COLLECTIONS_LIST);

    for (let i = 0; i < items.length; i++) {
      const itemId = (items[i] as Record<string, unknown>).Id as number;
      list.items.getById(itemId).update({ CollectionName: newName });
    }

    await execute();
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
