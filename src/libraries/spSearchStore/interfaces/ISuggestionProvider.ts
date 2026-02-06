import { ISuggestion } from './ISearchScope';
import { ISearchScope } from './ISearchScope';

/**
 * Context passed to suggestion and action providers.
 */
export interface ISearchContext {
  searchContextId: string;
  siteUrl: string;
  scope: ISearchScope;
}

/**
 * Suggestion provider — queried in parallel by the Search Box
 * to populate the suggestions dropdown.
 *
 * Built-in: RecentSearchProvider, TrendingQueryProvider, ManagedPropertyProvider
 */
export interface ISuggestionProvider {
  id: string;
  /** Section label in dropdown, e.g. "Recent", "Trending" */
  displayName: string;
  /** Sort order in dropdown — lower = higher priority */
  priority: number;
  maxResults: number;
  getSuggestions: (query: string, context: ISearchContext) => Promise<ISuggestion[]>;
  isEnabled: (context: ISearchContext) => boolean;
}
