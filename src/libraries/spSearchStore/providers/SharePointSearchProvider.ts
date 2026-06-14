import 'spfx-toolkit/lib/utilities/context/pnpImports/search';
import type { ISearchQuery as IPnPSearchQuery, ISearchResult as IPnPSearchResult, SearchResults } from '@pnp/sp/search';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { fetchManagedProperties, SearchService } from '@services/index';
import {
  ISearchDataProvider,
  ISearchQuery,
  ISearchResponse,
  IManagedProperty,
  ISearchResult,
  IActiveFilter,
  IFilterConfig,
  IRefiner,
  IRefinerValue,
  ISuggestion,
  IPromotedResultItem
} from '@interfaces/index';

/**
 * Default SharePoint Search data provider.
 * Wraps PnPjs sp.search() and maps results to the normalized interface.
 * Uses SPContext.sp for PnPjs access — no constructor injection needed.
 */
export class SharePointSearchProvider implements ISearchDataProvider {
  public readonly id: string = 'sharepoint-search';
  public readonly displayName: string = 'SharePoint Search';
  public readonly supportsRefiners: boolean = true;
  public readonly supportsCollapsing: boolean = true;
  public readonly supportsSorting: boolean = true;

  public async execute(query: ISearchQuery, signal: AbortSignal): Promise<ISearchResponse> {
    // Build PnPjs search request
    const startRow = (query.page - 1) * query.pageSize;

    // Apply scope KQL path restriction to query text
    let queryText = query.queryText || '*';

    // Normalize lowercase boolean operators to uppercase (SharePoint KQL requires AND/OR/NOT)
    queryText = this._normalizeKqlOperators(queryText);

    if (query.scope && query.scope.kqlPath) {
      queryText = queryText + ' ' + query.scope.kqlPath;
    }

    const searchRequest: IPnPSearchQuery = {
      Querytext: queryText,
      QueryTemplate: query.queryTemplate,
      RowLimit: query.pageSize,
      StartRow: startRow,
      SelectProperties: query.selectedProperties,
      TrimDuplicates: query.trimDuplicates !== undefined ? query.trimDuplicates : true,
      ClientType: 'SPSearch',
      EnableQueryRules: query.enableQueryRules !== undefined ? query.enableQueryRules : true,
    };

    // Add refiners
    if (query.refiners && query.refiners.length > 0) {
      searchRequest.Refiners = query.refiners.join(',');
    }

    // MISS-002 — Build refinement filters via the unified SearchService method.
    // Quotes the filter values first (FQL-aware), then delegates the cross-
    // property AND/OR combinator decision to `SearchService.buildRefinementFilters`.
    const quotedFilters = (query.filters || []).map(function quoteFilter(f: IActiveFilter): IActiveFilter {
      return { ...f, value: SharePointSearchProvider._quoteTokenValueStatic(f.value) };
    });
    const refinementFilters = SearchService.buildRefinementFilters(quotedFilters, query.operatorBetweenFilters || 'AND');

    // Merge persistent admin-configured refinement filters
    if (query.refinementFilters) {
      const persistent = query.refinementFilters.split(',').map(function (f: string): string { return f.trim(); }).filter(Boolean);
      for (let i = 0; i < persistent.length; i++) {
        refinementFilters.push(persistent[i]);
      }
    }

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

    // Add result source — explicit resultSourceId takes priority over scope
    if (query.resultSourceId) {
      searchRequest.SourceId = query.resultSourceId;
    } else if (query.scope && query.scope.resultSourceId) {
      searchRequest.SourceId = query.scope.resultSourceId;
    }

    // Add collapse specification with validation
    if (query.collapseSpecification) {
      searchRequest.CollapseSpecification = query.collapseSpecification;
    }

    // Check abort before making the API call
    if (signal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }

    // Guard: ensure SPContext is initialized (may not be ready if library bundle loaded first)
    if (!SPContext.isReady()) {
      throw new Error('SPContext not initialized — search provider waiting for web part initialization');
    }

    // Execute search via PnPjs
    const searchResults: SearchResults = await SPContext.sp.search(searchRequest);

    // Check abort after API call
    if (signal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }

    // Map results
    const items = this._mapResults(searchResults.PrimarySearchResults);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rawResults = searchResults.RawSearchResults as any;
    const refiners = this._mapRefiners(
      rawResults?.PrimaryQueryResult?.RefinementResults?.Refiners,
      query.filterConfig || []
    );
    const promotedResults = this._mapPromotedResults(rawResults?.PrimaryQueryResult?.SpecialTermResults?.Results);
    // SpellingSuggestion is the proper "Did you mean" field.
    // QueryModification contains internal query rule rewrites (e.g. "* -ContentClass=urn:...")
    // which are NOT user-facing spelling corrections.
    const querySuggestion = (rawResults?.PrimaryQueryResult?.RelevantResults?.Properties?.find(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (p: any) => p.Key === 'SpellingSuggestion'
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
      const results = await SPContext.sp.search({
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
    const result = await fetchManagedProperties();
    return result.properties;
  }

  /**
   * Normalize standalone KQL boolean operators to uppercase.
   * SharePoint Search requires AND/OR/NOT to be uppercase — lowercase
   * variants are treated as literal text. Only normalizes tokens that
   * appear as standalone words (preceded/followed by whitespace or parens),
   * NOT inside property:value pairs.
   */
  private _normalizeKqlOperators(queryText: string): string {
    if (!queryText) {
      return queryText;
    }
    // Match standalone and/or/not preceded by start-of-string, whitespace, or open-paren
    // and followed by whitespace, close-paren, or end-of-string.
    // Uses lookahead so consecutive operators don't consume shared boundaries.
    return queryText.replace(/(^|[\s(])(and|or|not)(?=[\s)]|$)/gi, function (_match: string, before: string, op: string): string {
      return before + op.toUpperCase();
    });
  }

  /**
   * Ensure a refinement token value is properly quoted for FQL.
   * FQL functions (range, string, and, or, not) pass through as-is.
   * Pre-quoted tokens ("...") pass through as-is.
   * Everything else gets wrapped in quotes — including ǂǂ hex tokens
   * and GP0|# taxonomy tokens that may arrive unquoted from PnPjs.
   *
   * Static so the `execute()` body can map filter values through it before
   * delegating to `SearchService.buildRefinementFilters` (MISS-002).
   */
  private static _quoteTokenValueStatic(value: string): string {
    // FQL functions — pass through as-is
    if (value.indexOf('range(') === 0 ||
      value.indexOf('string(') === 0 ||
      value.indexOf('and(') === 0 ||
      value.indexOf('or(') === 0 ||
      value.indexOf('not(') === 0) {
      return value;
    }
    // Already quoted — pass through as-is
    if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
      return value;
    }
    // All other values (hex tokens, taxonomy tokens, plain values) — quote them
    return '"' + value + '"';
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
   * Map raw refiner data to IRefiner[]. Delegates to
   * `mapRefinersWithPreprocessing` so the prefix-strip + delimiter-split
   * pass is exercised in the live search path AND unit-tested as a pure
   * helper.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _mapRefiners(rawRefiners: any[] | undefined, filterConfig: IFilterConfig[]): IRefiner[] {
    if (!rawRefiners || rawRefiners.length === 0) {
      return [];
    }
    return mapRefinersWithPreprocessing(rawRefiners, filterConfig);
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

// ─── Refiner preprocessing helpers ───────────────────────────────────────────

/**
 * Strip "type;#" prefix when appropriate and split values on a configured
 * delimiter. Exported so it can be unit-tested as a pure helper; consumed
 * by `SharePointSearchProvider._mapRefiners` in the live search path.
 *
 * Behaviour:
 *  - The cleaned/split label is returned as `name`.
 *  - The ORIGINAL raw refinement token (RefinementToken/RefinementValue) is
 *    preserved in `value`, so the eventual KQL/FQL refinement clause still
 *    matches SharePoint's multi-value serialization (`string;#Actual Value`).
 *  - Cleaned empty strings render as "(blank)" without losing the raw token.
 *  - When a delimiter is configured, per-entry counts are aggregated per
 *    cleaned token (dedupe + sum).
 */
export function mapRefinersWithPreprocessing(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  refinerResponses: any[],
  filterConfig: IFilterConfig[]
): IRefiner[] {
  const configByName = new Map<string, IFilterConfig>();
  for (let i = 0; i < filterConfig.length; i++) {
    configByName.set(filterConfig[i].managedProperty, filterConfig[i]);
  }

  return refinerResponses.map(function (r): IRefiner {
    const config = configByName.get(r.Name || '');
    const valueAggregator = new Map<string, { count: number; rawValue: string }>();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const entries: any[] = r.Entries || [];

    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i];
      const rawSource: string =
        (entry.RefinementName as string) || (entry.RefinementValue as string) || '';
      const rawToken: string =
        (entry.RefinementToken as string) || (entry.RefinementValue as string) || '';
      const count: number = parseInt(entry.RefinementCount || '0', 10) || 0;

      const tokens = splitEntryValue(rawSource, config ? config.valueSplitDelimiter : undefined);

      for (let j = 0; j < tokens.length; j++) {
        const cleaned = preprocessValue(tokens[j], config ? config.dataType : undefined);
        const displayName = cleaned.length === 0 ? '(blank)' : cleaned;
        const existing = valueAggregator.get(displayName);
        if (existing) {
          existing.count += count;
        } else {
          valueAggregator.set(displayName, { count: count, rawValue: rawToken });
        }
      }
    }

    const values: IRefinerValue[] = [];
    valueAggregator.forEach(function (agg, name): void {
      values.push({ name: name, value: agg.rawValue, count: agg.count, isSelected: false });
    });

    return { filterName: r.Name || '', values: values };
  });
}

function splitEntryValue(raw: string, delimiter: string | undefined): string[] {
  if (!delimiter) {
    return [raw];
  }
  const parts = raw.split(delimiter);
  const trimmed: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const t = parts[i].trim();
    if (t.length > 0) {
      trimmed.push(t);
    }
  }
  return trimmed.length > 0 ? trimmed : [raw];
}

/** Values stripped when dataType implicates SharePoint's "type;#value" serialization. */
const STRIP_DATA_TYPES = new Set<string>([
  'choiceMulti',
  'lookup',
  'calculated',
  'datetime',
  'number',
  'yesno',
]);

function preprocessValue(raw: string, dataType: IFilterConfig['dataType']): string {
  // Explicit opt-out: never strip when the admin says it's plain text.
  if (dataType === 'text') {
    return raw;
  }
  // Strip when explicitly one of the strip types, OR when dataType is
  // unspecified / 'auto' (heuristic mode).
  const shouldStrip =
    (dataType !== undefined && STRIP_DATA_TYPES.has(dataType)) ||
    dataType === undefined ||
    dataType === 'auto';

  if (!shouldStrip) {
    return raw;
  }

  // Heuristic: SharePoint prefix shape is `^[A-Za-z]+;#`.
  const m = /^[A-Za-z]+;#(.*)$/.exec(raw);
  if (m) {
    return m[1];
  }
  return raw;
}
