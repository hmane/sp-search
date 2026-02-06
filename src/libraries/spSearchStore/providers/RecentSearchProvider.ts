import { ISuggestionProvider, ISearchContext, ISuggestion } from '@interfaces/index';
import { SearchManagerService } from '@services/index';

/**
 * RecentSearchProvider — ISuggestionProvider that queries the SearchHistory
 * list for the current user's recent searches and returns them as suggestions.
 *
 * Registered in the SuggestionProviderRegistry during web part onInit.
 */
export class RecentSearchProvider implements ISuggestionProvider {
  public readonly id: string = 'recent-searches';
  public readonly displayName: string = 'Recent';
  public readonly priority: number = 10;
  public readonly maxResults: number = 5;

  private readonly _service: SearchManagerService;

  public constructor(service: SearchManagerService) {
    this._service = service;
  }

  public isEnabled(_context: ISearchContext): boolean {
    return true;
  }

  public async getSuggestions(query: string, _context: ISearchContext): Promise<ISuggestion[]> {
    try {
      // Load recent history entries (limited to maxResults)
      const history = await this._service.loadHistory(this.maxResults);

      if (history.length === 0) {
        return [];
      }

      const suggestions: ISuggestion[] = [];
      const normalizedQuery = query.toLowerCase().trim();

      for (let i = 0; i < history.length; i++) {
        const entry = history[i];

        // If the user has typed something, filter to matching entries
        if (normalizedQuery.length > 0) {
          const entryText = entry.queryText.toLowerCase();
          if (entryText.indexOf(normalizedQuery) < 0) {
            continue;
          }
        }

        suggestions.push({
          displayText: entry.queryText,
          groupName: 'Recent',
          iconName: 'History',
        });

        if (suggestions.length >= this.maxResults) {
          break;
        }
      }

      return suggestions;
    } catch {
      return [];
    }
  }
}
