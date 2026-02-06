import { ISuggestionProvider, ISearchContext, ISuggestion } from '@interfaces/index';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';

/**
 * TrendingQueryProvider — ISuggestionProvider that queries the SearchHistory
 * list for the current user's most frequently used queries.
 *
 * Fetches the top N most-frequently-used query terms (by current user) and
 * returns them as suggestions. Results are cached for 5 minutes to avoid
 * excessive list queries.
 *
 * NOTE: Author filter is required because SearchHistory WILL exceed 5,000 items.
 * Org-wide trending would require a pre-aggregated list or Flow-based approach.
 *
 * Requirement: §4.4.2
 */
export class TrendingQueryProvider implements ISuggestionProvider {
  public readonly id: string = 'trending-queries';
  public readonly displayName: string = 'Trending';
  public readonly priority: number = 20;
  public readonly maxResults: number = 5;

  /** Cached trending queries */
  private _cache: ISuggestion[] = [];
  /** Timestamp of last cache refresh */
  private _cacheTimestamp: number = 0;
  /** Cache TTL in milliseconds (5 minutes) */
  private static readonly CACHE_TTL: number = 5 * 60 * 1000;
  /** Whether a fetch is currently in progress */
  private _isFetching: boolean = false;

  public isEnabled(_context: ISearchContext): boolean {
    return true;
  }

  public async getSuggestions(query: string, _context: ISearchContext): Promise<ISuggestion[]> {
    try {
      // Refresh cache if stale
      const now = Date.now();
      if (now - this._cacheTimestamp > TrendingQueryProvider.CACHE_TTL && !this._isFetching) {
        await this._refreshCache();
      }

      if (this._cache.length === 0) {
        return [];
      }

      // Filter cached entries by current query input
      const normalizedQuery = query.toLowerCase().trim();
      if (normalizedQuery.length === 0) {
        return this._cache.slice(0, this.maxResults);
      }

      const filtered: ISuggestion[] = [];
      for (let i = 0; i < this._cache.length; i++) {
        const entry = this._cache[i];
        if (entry.displayText.toLowerCase().indexOf(normalizedQuery) >= 0) {
          filtered.push(entry);
          if (filtered.length >= this.maxResults) {
            break;
          }
        }
      }
      return filtered;
    } catch {
      return [];
    }
  }

  /**
   * Refresh the trending queries cache by querying SearchHistory.
   * Aggregates query text across all users and ranks by frequency.
   */
  private async _refreshCache(): Promise<void> {
    this._isFetching = true;
    try {
      // Query the SearchHistory list for recent entries (last 7 days)
      const sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
      const isoDate = sevenDaysAgo.toISOString();

      // CRITICAL: Always filter by Author first to avoid list view threshold (>5000 items).
      // SearchHistory is indexed on Author — filtering by current user prevents throttling.
      const items = await SPContext.sp.web.lists.getByTitle('SearchHistory').items
        .select('QueryText')
        .filter('Author eq ' + SPContext.currentUser.id + ' and Created ge datetime\'' + isoDate + '\'')
        .top(500)
        ();

      // Aggregate by query text (case-insensitive)
      const countMap: Map<string, { text: string; count: number }> = new Map();
      for (let i = 0; i < items.length; i++) {
        const text: string = items[i].QueryText;
        if (!text || text.trim().length === 0) {
          continue;
        }
        const key = text.toLowerCase().trim();
        const existing = countMap.get(key);
        if (existing) {
          existing.count++;
        } else {
          countMap.set(key, { text: text.trim(), count: 1 });
        }
      }

      // Sort by frequency (descending) and take top results
      const sorted = Array.from(countMap.values())
        .sort(function (a, b): number { return b.count - a.count; });

      const suggestions: ISuggestion[] = [];
      const limit = Math.min(sorted.length, 20); // Cache up to 20 for filtering
      for (let i = 0; i < limit; i++) {
        suggestions.push({
          displayText: sorted[i].text,
          groupName: 'Trending',
          iconName: 'Trending12',
        });
      }

      this._cache = suggestions;
      this._cacheTimestamp = Date.now();
    } catch {
      // On error, keep stale cache rather than clearing it
    } finally {
      this._isFetching = false;
    }
  }
}
