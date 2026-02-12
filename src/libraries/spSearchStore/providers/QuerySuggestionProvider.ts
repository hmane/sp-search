import 'spfx-toolkit/lib/utilities/context/pnpImports/search';
import type { ISuggestResult } from '@pnp/sp/search';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { ISuggestionProvider, ISearchContext, ISuggestion } from '@interfaces/index';

/**
 * QuerySuggestionProvider — ISuggestionProvider that calls the SharePoint
 * Search Suggestions API (`sp.searchSuggest()`) for real-time query
 * autocomplete as the user types.
 *
 * SharePoint returns completions based on the search index and what other
 * users have searched — this is the primary "autocomplete" experience.
 *
 * Priority 5 (highest) because these are the most contextually relevant
 * suggestions, appearing above Recent and Trending.
 *
 * Requirement: §4.4.2
 */
export class QuerySuggestionProvider implements ISuggestionProvider {
  public readonly id: string = 'query-suggestions';
  public readonly displayName: string = 'Suggestions';
  public readonly priority: number = 5;
  public readonly maxResults: number = 5;

  /** Simple LRU-like cache to avoid duplicate API calls for the same prefix */
  private _cache: Map<string, ISuggestion[]> = new Map();
  private static readonly MAX_CACHE_SIZE: number = 50;

  public isEnabled(_context: ISearchContext): boolean {
    return true;
  }

  public async getSuggestions(query: string, _context: ISearchContext): Promise<ISuggestion[]> {
    const trimmed = query.trim();
    if (trimmed.length < 2) {
      return [];
    }

    // Check cache first
    const cacheKey = trimmed.toLowerCase();
    const cached = this._cache.get(cacheKey);
    if (cached) {
      return cached;
    }

    try {
      if (!SPContext.isReady()) {
        return [];
      }

      const result: ISuggestResult = await SPContext.sp.searchSuggest({
        querytext: trimmed,
        count: this.maxResults,
        preQuery: true,
        hitHighlighting: false,
        capitalize: true,
        includePeople: false,
        queryRules: true,
        prefixMatch: true,
      });

      const suggestions: ISuggestion[] = [];

      // Parse query suggestions
      if (result.Queries && result.Queries.length > 0) {
        for (let i = 0; i < result.Queries.length && suggestions.length < this.maxResults; i++) {
          const entry = result.Queries[i];
          // PnPjs returns Queries as { Query: string } objects or raw strings
          const text: string = typeof entry === 'string' ? entry : (entry.Query || '');
          if (text.length > 0) {
            suggestions.push({
              displayText: text,
              groupName: 'Suggestions',
              iconName: 'Search',
            });
          }
        }
      }

      // Parse people suggestions as bonus results
      if (result.PeopleNames && result.PeopleNames.length > 0) {
        for (let i = 0; i < result.PeopleNames.length && suggestions.length < this.maxResults; i++) {
          const name: string = result.PeopleNames[i];
          if (name.length > 0) {
            suggestions.push({
              displayText: name,
              groupName: 'People',
              iconName: 'Contact',
            });
          }
        }
      }

      // Update cache (evict oldest if over limit)
      if (this._cache.size >= QuerySuggestionProvider.MAX_CACHE_SIZE) {
        const firstKey = this._cache.keys().next().value;
        if (firstKey !== undefined) {
          this._cache.delete(firstKey);
        }
      }
      this._cache.set(cacheKey, suggestions);

      return suggestions;
    } catch {
      // Swallow suggestion errors — non-critical
      return [];
    }
  }
}
