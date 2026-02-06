import { SPFI } from '@pnp/sp';
import '@pnp/sp/search';
import { ISearchQuery as IPnPSearchQuery, ISearchResult as IPnPSearchResult, SearchResults } from '@pnp/sp/search';
import {
  ISearchDataProvider,
  ISearchQuery,
  ISearchResponse,
  IManagedProperty,
  ISearchResult,
  IRefiner,
  IRefinerValue,
  ISuggestion,
  IPromotedResultItem
} from '@interfaces/index';

/**
 * Default SharePoint Search data provider.
 * Wraps PnPjs sp.search() and maps results to the normalized interface.
 */
export class SharePointSearchProvider implements ISearchDataProvider {
  public readonly id: string = 'sharepoint-search';
  public readonly displayName: string = 'SharePoint Search';
  public readonly supportsRefiners: boolean = true;
  public readonly supportsCollapsing: boolean = true;
  public readonly supportsSorting: boolean = true;

  private readonly _sp: SPFI;

  public constructor(sp: SPFI) {
    this._sp = sp;
  }

  public async execute(query: ISearchQuery, signal: AbortSignal): Promise<ISearchResponse> {
    // Build PnPjs search request
    const startRow = (query.page - 1) * query.pageSize;

    const searchRequest: IPnPSearchQuery = {
      Querytext: query.queryText || '*',
      QueryTemplate: query.queryTemplate,
      RowLimit: query.pageSize,
      StartRow: startRow,
      SelectProperties: query.selectedProperties,
      TrimDuplicates: query.trimDuplicates !== undefined ? query.trimDuplicates : true,
      ClientType: 'SPSearch',
      EnableQueryRules: true,
    };

    // Add refiners
    if (query.refiners && query.refiners.length > 0) {
      searchRequest.Refiners = query.refiners.join(',');
    }

    // Add refinement filters (active filter values)
    const refinementFilters = this._buildRefinementFilters(query.filters);
    if (refinementFilters.length > 0) {
      searchRequest.RefinementFilters = refinementFilters;
    }

    // Add sort
    if (query.sort) {
      searchRequest.SortList = [{
        Property: query.sort.property,
        Direction: query.sort.direction === 'Ascending' ? 0 : 1,
      }];
    }

    // Add result source
    if (query.resultSourceId) {
      searchRequest.SourceId = query.resultSourceId;
    }

    // Add collapse specification with validation
    if (query.collapseSpecification) {
      searchRequest.CollapseSpecification = query.collapseSpecification;
    }

    // Check abort before making the API call
    if (signal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }

    // Execute search
    const searchResults: SearchResults = await this._sp.search(searchRequest);

    // Check abort after API call
    if (signal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }

    // Map results
    const items = this._mapResults(searchResults.PrimarySearchResults);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rawResults = searchResults.RawSearchResults as any;
    const refiners = this._mapRefiners(rawResults?.PrimaryQueryResult?.RefinementResults?.Refiners);
    const promotedResults = this._mapPromotedResults(rawResults?.PrimaryQueryResult?.SpecialTermResults?.Results);
    const querySuggestion = (rawResults?.PrimaryQueryResult?.RelevantResults?.Properties?.find(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (p: any) => p.Key === 'QueryModification'
    )?.Value as string) || undefined;

    return {
      items,
      totalCount: searchResults.TotalRows,
      refiners,
      promotedResults,
      querySuggestion,
    };
  }

  public async getSuggestions(query: string, signal: AbortSignal): Promise<ISuggestion[]> {
    if (!query || query.length < 2) {
      return [];
    }

    if (signal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }

    try {
      const results = await this._sp.search({
        Querytext: query,
        RowLimit: 0,
        QueryTemplate: '{searchTerms}',
        ClientType: 'SPSearch',
        EnableQueryRules: true,
      });

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const suggestion = (results as any).SpellingSuggestion as string | undefined;
      if (suggestion) {
        return [{
          displayText: suggestion,
          groupName: 'Suggestions',
        }];
      }
    } catch {
      // Swallow suggestion errors — non-critical
    }

    return [];
  }

  public async getSchema(): Promise<IManagedProperty[]> {
    // Fetch managed properties from the search schema
    // This uses the search REST API to get property metadata
    try {
      const results = await this._sp.search({
        Querytext: '*',
        RowLimit: 1,
        SelectProperties: ['Title'],
        ClientType: 'SPSearch',
      });

      // Extract properties from the result metadata
      const properties: IManagedProperty[] = [];
      const rawProperties = results.RawSearchResults?.PrimaryQueryResult?.RelevantResults?.Properties;
      if (rawProperties) {
        for (const prop of rawProperties) {
          if (prop.Key && prop.Key !== 'RowCount' && prop.Key !== 'TotalRows') {
            properties.push({
              name: prop.Key,
              type: typeof prop.Value === 'number' ? 'Integer' : 'Text',
              queryable: true,
              retrievable: true,
              refinable: false,
              sortable: false,
            });
          }
        }
      }

      return properties;
    } catch {
      return [];
    }
  }

  /**
   * Build refinement filter strings from active filters.
   * Groups by filterName, combines values with or() within groups.
   */
  private _buildRefinementFilters(filters: ISearchQuery['filters']): string[] {
    if (!filters || filters.length === 0) {
      return [];
    }

    // Group filters by filterName
    const grouped: Record<string, string[]> = {};
    for (const filter of filters) {
      if (!grouped[filter.filterName]) {
        grouped[filter.filterName] = [];
      }
      grouped[filter.filterName].push(filter.value);
    }

    // Build refinement filter strings
    const result: string[] = [];
    const keys = Object.keys(grouped);
    for (const key of keys) {
      const values = grouped[key];
      if (values.length === 1) {
        // Check if value is already FQL (range, etc.)
        const val = values[0];
        if (this._isFqlExpression(val)) {
          result.push(key + ':' + val);
        } else {
          result.push(key + ':"ǂǂ' + this._encodeRefinementValue(val) + '"');
        }
      } else {
        // Multiple values — combine with or()
        const encoded = values.map((v) => {
          if (this._isFqlExpression(v)) {
            return v;
          }
          return '"ǂǂ' + this._encodeRefinementValue(v) + '"';
        });
        result.push(key + ':or(' + encoded.join(',') + ')');
      }
    }

    return result;
  }

  /**
   * Check if a value is already an FQL expression (range, and, or, etc.)
   */
  private _isFqlExpression(value: string): boolean {
    return value.indexOf('range(') === 0 ||
      value.indexOf('and(') === 0 ||
      value.indexOf('or(') === 0 ||
      value.indexOf('not(') === 0;
  }

  /**
   * Hex-encode a refinement value for the ǂǂ prefix format.
   */
  private _encodeRefinementValue(value: string): string {
    let hex = '';
    for (let i = 0; i < value.length; i++) {
      const code = value.charCodeAt(i).toString(16);
      // Pad to 4 chars (ES5-safe alternative to padStart)
      hex += ('0000' + code).slice(-4);
    }
    return hex;
  }

  /**
   * Map raw PnP search results to normalized ISearchResult[].
   */
  private _mapResults(rawResults: IPnPSearchResult[]): ISearchResult[] {
    if (!rawResults) {
      return [];
    }

    return rawResults.map((raw): ISearchResult => {
      const properties: Record<string, unknown> = {};
      // Copy all raw properties
      const keys = Object.keys(raw);
      for (const key of keys) {
        properties[key] = (raw as Record<string, unknown>)[key];
      }

      return {
        key: (raw as Record<string, unknown>).DocId as string ||
          (raw as Record<string, unknown>).UniqueId as string ||
          raw.Path || '',
        title: raw.Title || '',
        url: raw.Path || '',
        summary: this._sanitizeHighlighting(
          (raw as Record<string, unknown>).HitHighlightedSummary as string || ''
        ),
        author: {
          displayText: raw.Author || '',
          email: (raw as Record<string, unknown>).AuthorOWSUSER as string || '',
          imageUrl: undefined,
        },
        created: (raw as Record<string, unknown>).Created as string || '',
        modified: String(raw.LastModifiedTime || ''),
        fileType: (raw as Record<string, unknown>).FileType as string ||
          (raw as Record<string, unknown>).FileExtension as string || '',
        fileSize: parseInt((raw as Record<string, unknown>).Size as string || '0', 10) || 0,
        siteName: (raw as Record<string, unknown>).SiteTitle as string ||
          (raw as Record<string, unknown>).SiteName as string || '',
        siteUrl: (raw as Record<string, unknown>).SPSiteURL as string || '',
        thumbnailUrl: (raw as Record<string, unknown>).PictureThumbnailURL as string ||
          (raw as Record<string, unknown>).ServerRedirectedPreviewURL as string || '',
        properties,
      };
    });
  }

  /**
   * Clean up SharePoint hit-highlighting tags.
   * Replaces <ddd/> with ... and <c0>...</c0> with <mark>...</mark>
   */
  private _sanitizeHighlighting(summary: string): string {
    if (!summary) {
      return '';
    }
    return summary
      .replace(/<ddd\/>/g, '...')
      .replace(/<c0>/g, '<mark>')
      .replace(/<\/c0>/g, '</mark>');
  }

  /**
   * Map raw refiner data to IRefiner[].
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _mapRefiners(rawRefiners: any[] | undefined): IRefiner[] {
    if (!rawRefiners || rawRefiners.length === 0) {
      return [];
    }

    return rawRefiners.map((rawRefiner): IRefiner => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const entries: any[] = rawRefiner.Entries || [];
      const values: IRefinerValue[] = entries.map((entry): IRefinerValue => ({
        name: entry.RefinementName || entry.RefinementValue || '',
        value: entry.RefinementToken || entry.RefinementValue || '',
        count: parseInt(entry.RefinementCount || '0', 10) || 0,
        isSelected: false,
      }));

      return {
        filterName: rawRefiner.Name || '',
        values,
      };
    });
  }

  /**
   * Map promoted/best bet results.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _mapPromotedResults(rawPromoted: any[] | undefined): IPromotedResultItem[] {
    if (!rawPromoted || rawPromoted.length === 0) {
      return [];
    }

    return rawPromoted.map((raw): IPromotedResultItem => ({
      title: raw.Title || '',
      url: raw.Url || '',
      description: raw.Description || undefined,
    }));
  }
}
