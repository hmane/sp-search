import { ISearchDataProvider, ISearchContext, ISearchQuery, ISearchResult, ISuggestion, ISuggestionProvider } from '@interfaces/index';
import type { IRegistry } from '@interfaces/index';

const QUICK_RESULT_PROPERTIES = ['Title', 'Path', 'SiteName', 'FileType'];

function buildSecondaryText(item: ISearchResult): string {
  const parts: string[] = [];

  if (item.siteName) {
    parts.push(item.siteName);
  }

  if (item.fileType) {
    parts.push(item.fileType.toUpperCase());
  }

  return parts.join(' • ') || item.url;
}

export class QuickResultsSuggestionProvider implements ISuggestionProvider {
  public readonly id: string = 'quick-results';
  public readonly displayName: string = 'Quick Results';
  public readonly priority: number = 12;
  public readonly maxResults: number = 3;

  private readonly _dataProviderRegistry: IRegistry<ISearchDataProvider>;

  public constructor(dataProviderRegistry: IRegistry<ISearchDataProvider>) {
    this._dataProviderRegistry = dataProviderRegistry;
  }

  public isEnabled(_context: ISearchContext): boolean {
    return true;
  }

  public async getSuggestions(query: string, context: ISearchContext): Promise<ISuggestion[]> {
    const trimmed = query.trim();
    if (trimmed.length < 2) {
      return [];
    }

    const provider = this._dataProviderRegistry.get('sharepoint-search') || this._dataProviderRegistry.getAll()[0];
    if (!provider) {
      return [];
    }

    const quickQuery: ISearchQuery = {
      queryText: trimmed,
      queryTemplate: '{searchTerms}',
      scope: context.scope,
      filters: [],
      sort: undefined,
      page: 1,
      pageSize: this.maxResults,
      selectedProperties: QUICK_RESULT_PROPERTIES,
      refiners: [],
      trimDuplicates: true,
      enableQueryRules: true,
      operatorBetweenFilters: 'AND'
    };

    try {
      const response = await provider.execute(quickQuery, new AbortController().signal);
      const suggestions: ISuggestion[] = [];

      for (let i = 0; i < response.items.length && suggestions.length < this.maxResults; i++) {
        const item = response.items[i];

        if (!item.title || !item.url) {
          continue;
        }

        suggestions.push({
          displayText: item.title,
          secondaryText: buildSecondaryText(item),
          groupName: 'Quick Results',
          filePath: item.url,
          iconName: 'Page',
          action: function (): void {
            window.location.href = item.url;
          }
        });
      }

      return suggestions;
    } catch {
      return [];
    }
  }
}
