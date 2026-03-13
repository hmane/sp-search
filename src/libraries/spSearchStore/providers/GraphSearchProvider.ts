import type { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import {
  ISearchDataProvider,
  ISearchQuery,
  ISearchResponse,
  ISearchResult,
  IRefiner,
  IRefinerValue
} from '@interfaces/index';

// ─── Provider configuration ─────────────────────────────────────────────────

export interface IGraphSearchProviderConfig {
  /**
   * Provider ID. Override when registering multiple Graph providers
   * (e.g. one for files and one for people on the same page).
   * Default: 'graph-search'.
   */
  id?: string;
  /**
   * Microsoft Graph entity types to target.
   * Default: ['driveItem', 'listItem', 'site'].
   * Use ['person'] for a People vertical; ['externalItem'] for connectors.
   */
  entityTypes?: string[];
  /**
   * Required only when entityTypes includes 'externalItem'.
   * Format: ['/external/connections/{connectionId}'].
   */
  contentSources?: string[];
}

// ─── Internal Graph API shapes ───────────────────────────────────────────────

interface IGraphSearchAggregation {
  field: string;
  size: number;
  bucketDefinition: {
    sortBy: string;
    isDescending: boolean;
    minimumCount: number;
  };
}

interface IGraphSearchAggregationBucket {
  key: string;
  count: number;
  aggregationFilterToken: string;
}

interface IGraphSearchAggregationResult {
  field: string;
  buckets: IGraphSearchAggregationBucket[];
}

interface IGraphSearchHit {
  _id: string;
  rank: number;
  summary?: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  resource: Record<string, any>;
}

interface IGraphSearchHitsContainer {
  hits: IGraphSearchHit[];
  total: number;
  moreResultsAvailable: boolean;
  aggregations?: IGraphSearchAggregationResult[];
}

// ─── Provider ────────────────────────────────────────────────────────────────

/**
 * Microsoft Graph Search data provider.
 *
 * Routes queries through the Graph `/search/query` endpoint instead of
 * SharePoint Search. Best suited for:
 * - People verticals (entityTypes: ['person'])
 * - External connector content (entityTypes: ['externalItem'])
 * - Cross-tenant / hybrid search scenarios
 *
 * Limitations vs. SharePointSearchProvider:
 * - Does not support CollapseSpecification (Graph has no equivalent)
 * - FQL refinement tokens (ǂǂ…, GP0|#…, range()) are not forwarded —
 *   Graph aggregation filters use a different format. B3 work adds
 *   Graph-specific filter formatters.
 * - QueryRules (best bets) are not available; promotedResults always empty.
 * - Spelling suggestions are not available; querySuggestion always undefined.
 *
 * Usage (in web part onInit):
 *   const graphClient = await this.context.msGraphClientFactory.getClient('3');
 *   store.getState().registries.dataProviders.register(
 *     new GraphSearchProvider(graphClient, { entityTypes: ['person'] })
 *   );
 */
export class GraphSearchProvider implements ISearchDataProvider {
  public readonly id: string;
  public readonly displayName: string;
  public readonly supportsRefiners: boolean;
  public readonly supportsCollapsing: boolean = false;
  public readonly supportsSorting: boolean;

  private readonly _graphClient: MSGraphClientV3;
  private readonly _entityTypes: string[];
  private readonly _contentSources: string[];

  public constructor(graphClient: MSGraphClientV3, config: IGraphSearchProviderConfig = {}) {
    this._graphClient = graphClient;
    this.id = config.id || 'graph-search';
    this._entityTypes = config.entityTypes || ['driveItem', 'listItem', 'site'];
    this._contentSources = config.contentSources || [];

    // Display name reflects entity scope for property pane clarity
    this.displayName = 'Microsoft Graph Search (' + this._entityTypes.join(', ') + ')';

    // Aggregations (refiners) work for file-shaped entity types
    const refinableTypes = ['driveItem', 'listItem', 'externalItem'];
    this.supportsRefiners = this._entityTypes.some((t) => refinableTypes.indexOf(t) >= 0);

    // Sorting is supported for SharePoint content entity types
    const sortableTypes = ['driveItem', 'listItem'];
    this.supportsSorting = this._entityTypes.some((t) => sortableTypes.indexOf(t) >= 0);
  }

  public async execute(query: ISearchQuery, signal: AbortSignal): Promise<ISearchResponse> {
    if (signal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const searchRequest: Record<string, any> = {
      entityTypes: this._entityTypes,
      query: {
        queryString: this._buildQueryString(query)
      },
      from: (query.page - 1) * query.pageSize,
      size: query.pageSize,
    };

    // Selected fields — Graph uses 'fields' to scope retrieved properties
    if (query.selectedProperties && query.selectedProperties.length > 0) {
      searchRequest.fields = query.selectedProperties;
    }

    // Sort properties — only added when the provider advertises sorting support
    if (query.sort && this.supportsSorting) {
      searchRequest.sortProperties = [{
        name: query.sort.property,
        isDescending: query.sort.direction === 'Descending'
      }];
    }

    // Aggregations — request refiner buckets from Graph
    if (query.refiners && query.refiners.length > 0 && this.supportsRefiners) {
      searchRequest.aggregations = query.refiners.map((refinerName): IGraphSearchAggregation => ({
        field: refinerName,
        size: 10,
        bucketDefinition: {
          sortBy: 'count',
          isDescending: true,
          minimumCount: 0
        }
      }));
    }

    // Active filter aggregation filters — plain equality filters only.
    // FQL-specific tokens (range, ǂǂ, GP0|#) are SharePoint-format and cannot
    // be forwarded to Graph. Those require Graph-specific formatters (B3).
    const aggFilters = this._buildAggregationFilters(query.filters);
    if (aggFilters.length > 0) {
      searchRequest.aggregationFilters = aggFilters;
    }

    // Content sources — required for externalItem entity type
    if (this._contentSources.length > 0) {
      searchRequest.contentSources = this._contentSources;
    }

    // Execute Graph Search POST request
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const response: any = await this._graphClient
      .api('/search/query')
      .post({ requests: [searchRequest] });

    // Check abort after the (non-cancellable) HTTP call returns
    if (signal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }

    const container: IGraphSearchHitsContainer | undefined = response?.value?.[0]?.hitsContainers?.[0];

    if (!container) {
      return { items: [], totalCount: 0, refiners: [], promotedResults: [], querySuggestion: undefined };
    }

    return {
      items: this._mapHits(container.hits || []),
      totalCount: container.total || 0,
      refiners: this._mapAggregations(container.aggregations || []),
      promotedResults: [],
      querySuggestion: undefined
    };
  }

  // ─── Private helpers ───────────────────────────────────────────────────────

  /**
   * Build the Graph queryString from ISearchQuery.
   *
   * Graph accepts KQL in the queryString, so scope path restrictions
   * (`path:https://...`) and queryTemplate substitution work the same
   * way as in SharePoint Search.
   */
  private _buildQueryString(query: ISearchQuery): string {
    let qs = query.queryText || '*';

    // Apply query template: substitute {searchTerms} placeholder.
    // Templates that are just the placeholder (the default) collapse to
    // the raw query text without modification.
    const template = (query.queryTemplate || '').trim();
    if (template && template !== '{searchTerms}') {
      qs = template.replace(/\{searchTerms\}/gi, qs);
    }

    // Scope restriction — append the KQL path predicate
    if (query.scope && query.scope.kqlPath) {
      qs = qs + ' ' + query.scope.kqlPath;
    }

    return qs.trim() || '*';
  }

  /**
   * Convert active equality filters to Graph aggregation filter format.
   *
   * Only handles plain-value equality filters. FQL-specific tokens used
   * by SharePoint Search (range(), ǂǂ…, GP0|#…, string()) are silently
   * skipped because Graph does not understand them.
   */
  private _buildAggregationFilters(filters: ISearchQuery['filters']): string[] {
    if (!filters || filters.length === 0) {
      return [];
    }

    // Group by filterName (same behavior as SharePointSearchProvider)
    const grouped: Record<string, string[]> = {};
    for (const filter of filters) {
      // Skip FQL-specific tokens
      const v = filter.value;
      if (v.indexOf('range(') === 0 ||
          v.indexOf('string(') === 0 ||
          v.indexOf('and(') === 0 ||
          v.indexOf('or(') === 0 ||
          v.indexOf('not(') === 0 ||
          v.indexOf('ǂǂ') === 0 ||
          v.indexOf('GP0|#') === 0) {
        continue;
      }

      if (!grouped[filter.filterName]) {
        grouped[filter.filterName] = [];
      }
      grouped[filter.filterName].push(v);
    }

    const result: string[] = [];
    const keys = Object.keys(grouped);
    for (const key of keys) {
      const values = grouped[key];
      // Graph format: "field:equals(\"value\")" or "field:or(equals(\"a\"),equals(\"b\"))"
      if (values.length === 1) {
        result.push(key + ':equals("' + values[0].replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '")');
      } else {
        const parts = values.map((v) => 'equals("' + v.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '")');
        result.push(key + ':or(' + parts.join(',') + ')');
      }
    }

    return result;
  }

  /**
   * Map Graph search hits to normalized ISearchResult[].
   * Handles driveItem, listItem, site, and person entity types.
   */
  private _mapHits(hits: IGraphSearchHit[]): ISearchResult[] {
    if (!hits) {
      return [];
    }

    return hits.map((hit): ISearchResult => {
      const resource = hit.resource || {};
      const odataType: string = resource['@odata.type'] || '';

      if (odataType.indexOf('person') >= 0) {
        return this._mapPersonHit(hit);
      }
      if (odataType.indexOf('listItem') >= 0) {
        return this._mapListItemHit(hit);
      }
      // driveItem, site, default
      return this._mapDriveItemHit(hit);
    });
  }

  private _mapDriveItemHit(hit: IGraphSearchHit): ISearchResult {
    const r = hit.resource;
    const createdBy = r.createdBy?.user || r.lastModifiedBy?.user || {};
    const parentRef = r.parentReference || {};

    return {
      key: hit._id || r.id || '',
      title: r.name || r.displayName || '',
      url: r.webUrl || '',
      summary: hit.summary || '',
      author: {
        displayText: createdBy.displayName || '',
        email: createdBy.email || createdBy.mail || createdBy.userPrincipalName || '',
        imageUrl: undefined
      },
      created: r.createdDateTime || '',
      modified: r.lastModifiedDateTime || '',
      fileType: this._getFileTypeFromName(r.name || ''),
      fileSize: typeof r.size === 'number' ? r.size : parseInt(String(r.size || '0'), 10) || 0,
      siteName: parentRef.name || r.parentReference?.siteId || '',
      siteUrl: '',
      thumbnailUrl: r.thumbnailUrl || '',
      properties: r as Record<string, unknown>
    };
  }

  private _mapListItemHit(hit: IGraphSearchHit): ISearchResult {
    const r = hit.resource;
    const fields = r.fields || {};
    const createdBy = r.createdBy?.user || {};

    return {
      key: hit._id || r.id || '',
      title: fields.Title || r.displayName || r.name || '',
      url: r.webUrl || '',
      summary: hit.summary || '',
      author: {
        displayText: createdBy.displayName || fields.Author || '',
        email: createdBy.email || createdBy.mail || createdBy.userPrincipalName || '',
        imageUrl: undefined
      },
      created: r.createdDateTime || fields.Created || '',
      modified: r.lastModifiedDateTime || fields.Modified || '',
      fileType: '',
      fileSize: 0,
      siteName: '',
      siteUrl: '',
      thumbnailUrl: '',
      properties: { ...fields, ...r } as Record<string, unknown>
    };
  }

  private _mapPersonHit(hit: IGraphSearchHit): ISearchResult {
    const r = hit.resource;
    // scoredEmailAddresses is an array; take the first one
    const primaryEmail: string =
      (Array.isArray(r.scoredEmailAddresses) && r.scoredEmailAddresses.length > 0
        ? r.scoredEmailAddresses[0].address
        : '') ||
      r.userPrincipalName || '';

    // Build a profile URL from the UPN — works for M365-connected tenants
    const upn: string = r.userPrincipalName || primaryEmail;
    const profileUrl = upn ? 'https://teams.microsoft.com/_#/profile/' + encodeURIComponent(upn) : '';

    return {
      key: hit._id || r.id || '',
      title: r.displayName || '',
      url: profileUrl,
      summary: hit.summary || [r.jobTitle, r.department, r.officeLocation].filter(Boolean).join(' · '),
      author: {
        displayText: r.displayName || '',
        email: primaryEmail,
        imageUrl: undefined
      },
      created: '',
      modified: '',
      fileType: '',
      fileSize: 0,
      siteName: r.department || '',
      siteUrl: '',
      thumbnailUrl: '',
      properties: r as Record<string, unknown>
    };
  }

  /**
   * Map Graph aggregation buckets to normalized IRefiner[].
   */
  private _mapAggregations(aggregations: IGraphSearchAggregationResult[]): IRefiner[] {
    if (!aggregations || aggregations.length === 0) {
      return [];
    }

    return aggregations.map((agg): IRefiner => {
      const values: IRefinerValue[] = (agg.buckets || []).map((bucket): IRefinerValue => ({
        name: bucket.key,
        // Store the aggregationFilterToken so the filter can be forwarded
        // back to Graph as-is when the user selects the value
        value: bucket.aggregationFilterToken || bucket.key,
        count: bucket.count || 0,
        isSelected: false
      }));

      return {
        filterName: agg.field,
        values
      };
    });
  }

  /**
   * Derive a file extension from a filename.
   * Returns empty string for items without an extension (folders, sites).
   */
  private _getFileTypeFromName(name: string): string {
    const lastDot = name.lastIndexOf('.');
    if (lastDot < 0 || lastDot === name.length - 1) {
      return '';
    }
    return name.substring(lastDot + 1).toLowerCase();
  }
}
