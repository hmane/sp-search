import type { IRegistry, ISuggestionProvider, ISearchDataProvider } from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import { RecentSearchProvider, TrendingQueryProvider, ManagedPropertyProvider } from '@providers/index';

/**
 * Register all built-in suggestion providers into the given SuggestionProviderRegistry.
 * Called once from SpSearchBoxWebPart.onInit() after initializeSearchContext.
 *
 * Providers are registered only if not already present (idempotent).
 *
 * Priority order (lower = higher):
 *   10 — Recent searches (user-specific)
 *   20 — Trending queries (org-wide)
 *   30 — Managed property suggestions
 */
export function registerBuiltInSuggestions(
  registry: IRegistry<ISuggestionProvider>,
  managerService: SearchManagerService,
  dataProviderRegistry: IRegistry<ISearchDataProvider>
): void {
  // Recent searches
  if (!registry.get('recent-searches')) {
    registry.register(new RecentSearchProvider(managerService));
  }

  // Trending queries (org-wide popular searches)
  if (!registry.get('trending-queries')) {
    registry.register(new TrendingQueryProvider());
  }

  // Managed property suggestions (KQL property:value)
  if (!registry.get('managed-property')) {
    registry.register(new ManagedPropertyProvider(dataProviderRegistry));
  }
}
