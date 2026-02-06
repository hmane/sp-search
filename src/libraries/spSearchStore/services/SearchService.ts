import { IActiveFilter, ISortField } from '@interfaces/index';
import { ITokenContext, TokenService } from './TokenService';

/**
 * Default managed properties requested from the SharePoint Search API.
 * Custom properties provided by the caller are merged on top of these.
 */
const DEFAULT_SELECTED_PROPERTIES: string[] = [
  'Title',
  'Path',
  'Filename',
  'Author',
  'AuthorOWSUSER',
  'Created',
  'LastModifiedTime',
  'FileType',
  'FileExtension',
  'SecondaryFileExtension',
  'contentclass',
  'HitHighlightedSummary',
  'HitHighlightedProperties',
  'SiteName',
  'SiteTitle',
  'SPSiteURL',
  'ServerRedirectedURL',
  'ServerRedirectedPreviewURL',
  'PictureThumbnailURL',
  'ParentLink',
  'ViewsLifeTime',
  'Size',
  'NormSiteID',
  'NormListID',
  'NormUniqueID',
  'DocId',
  'IsDocument',
];

/**
 * KQL query assembly and refinement token construction.
 *
 * All methods are static (no instance state needed). The service
 * coordinates with {@link TokenService} for template token resolution
 * and provides helpers for building the full search request payload.
 */
export class SearchService {

  /**
   * Build the final KQL query string from a template, user query text,
   * and active filter context.
   *
   * Steps:
   * 1. Resolve template tokens (e.g. `{searchTerms}`, `{Site.URL}`, etc.)
   *    using the provided {@link ITokenContext}.
   * 2. Trim the result.
   *
   * @param queryTemplate - KQL query template containing `{...}` tokens.
   * @param queryText     - The raw search query entered by the user.
   * @param activeFilters - Currently active filter selections (used for
   *                        refinement filters, built separately via
   *                        {@link buildRefinementFilters}).
   * @param tokenContext  - Context values for token resolution.
   * @returns The fully resolved KQL query string.
   */
  public static buildKqlQuery(
    queryTemplate: string,
    queryText: string,
    activeFilters: IActiveFilter[],
    tokenContext: ITokenContext
  ): string {
    // Ensure the token context carries the latest query text
    const context: ITokenContext = {
      ...tokenContext,
      queryText,
    };

    const resolvedQuery: string = TokenService.resolveTokens(queryTemplate, context);

    return resolvedQuery.trim();
  }

  /**
   * Build refinement filter tokens for the SharePoint Search REST API.
   *
   * Filters are grouped by `filterName` (managed property). Within each
   * group, multiple values are combined with `or(value1,value2,...)`.
   * Groups themselves are returned as separate array entries so that the
   * search API combines them with AND logic.
   *
   * If a filter value looks like FQL (starts with `range(` or contains
   * a comma), it is passed through as-is without wrapping.
   *
   * @example
   * ```
   * // Input:
   * [
   *   { filterName: 'FileType', value: '"docx"', operator: 'OR' },
   *   { filterName: 'FileType', value: '"pptx"', operator: 'OR' },
   *   { filterName: 'Author',   value: '"John"', operator: 'OR' },
   * ]
   * // Output:
   * [
   *   'FileType:or("docx","pptx")',
   *   'Author:"John"'
   * ]
   * ```
   *
   * @param activeFilters - The currently active filter selections.
   * @returns An array of refinement filter strings.
   */
  public static buildRefinementFilters(activeFilters: IActiveFilter[]): string[] {
    if (!activeFilters || activeFilters.length === 0) {
      return [];
    }

    // Group filter values by managed property name
    const grouped: Map<string, string[]> = new Map();

    for (const filter of activeFilters) {
      const existing: string[] | undefined = grouped.get(filter.filterName);
      if (existing !== undefined) {
        existing.push(filter.value);
      } else {
        grouped.set(filter.filterName, [filter.value]);
      }
    }

    const refinementFilters: string[] = [];

    grouped.forEach((values: string[], filterName: string) => {
      if (values.length === 1) {
        // Single value: emit directly (FQL or simple token)
        refinementFilters.push(`${filterName}:${values[0]}`);
      } else {
        // Multiple values: combine with or()
        refinementFilters.push(`${filterName}:or(${values.join(',')})`);
      }
    });

    return refinementFilters;
  }

  /**
   * Build the sort parameter for the SharePoint Search REST API.
   *
   * @param sort - The sort field configuration, or `undefined` for relevance
   *               (default) sorting.
   * @returns An array with a single sort entry, or an empty array if no
   *          explicit sort is defined.
   *
   * @example
   * ```
   * buildSortList({ property: 'LastModifiedTime', direction: 'Descending' })
   * // => [{ Property: 'LastModifiedTime', Direction: 1 }]
   * ```
   */
  public static buildSortList(
    sort: ISortField | undefined
  ): Array<{ Property: string; Direction: number }> {
    if (sort === undefined) {
      return [];
    }

    const direction: number = sort.direction === 'Ascending' ? 0 : 1;

    return [
      {
        Property: sort.property,
        Direction: direction,
      },
    ];
  }

  /**
   * Build the selected properties array for the search request.
   *
   * Merges the default set of managed properties with any custom
   * properties provided by the caller. Duplicates are removed.
   *
   * @param customProperties - Additional managed property names to include.
   * @returns A deduplicated array of property names.
   */
  public static buildSelectedProperties(
    customProperties?: string[]
  ): string[] {
    if (!customProperties || customProperties.length === 0) {
      return [...DEFAULT_SELECTED_PROPERTIES];
    }

    // Use a Set to deduplicate while preserving order (defaults first)
    const propertySet: Set<string> = new Set(DEFAULT_SELECTED_PROPERTIES);

    for (const prop of customProperties) {
      propertySet.add(prop);
    }

    return Array.from(propertySet);
  }

}
