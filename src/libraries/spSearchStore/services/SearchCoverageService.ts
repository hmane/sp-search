import 'spfx-toolkit/lib/utilities/context/pnpImports/lists';
import 'spfx-toolkit/lib/utilities/context/pnpImports/pages';
import 'spfx-toolkit/lib/utilities/context/pnpImports/search';
import { CacheNever } from '@pnp/queryable';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import type { ISearchQuery, ISearchScope } from '@interfaces/index';
import { SharePointSearchProvider } from '@providers/SharePointSearchProvider';

const COVERAGE_SCOPE: ISearchScope = {
  id: 'coverage',
  label: 'Coverage'
};

const DEFAULT_QUERY_TEMPLATE = '{searchTerms}';
const MAX_MISSING_SAMPLE_ITEMS = 8;
const MAX_MISSING_SAMPLE_SCAN = 16;
const SEARCH_RESULTS_WEBPART_ID = '1836671c-a710-45b4-9a83-55c65344a3d5';

export interface ICoverageProfile {
  id: string;
  title: string;
  description?: string;
  queryTemplate: string;
  resultSourceId?: string;
  sourceUrls: string[];
  contentTypeIds: string[];
  excludePaths: string[];
  includeFolders: boolean;
  trimDuplicates: boolean;
  refinementFilters?: string;
}

export interface ICoverageSourceResult {
  title: string;
  sourceUrl: string;
  sourceCount: number;
  searchCount: number;
  searchCountTrimmed: number;
  searchCountUntrimmed: number;
  duplicateDelta: number;
  delta: number;
  noCrawl: boolean;
  hidden: boolean;
}

export interface ICoverageMissingItem {
  title: string;
  path: string;
  sourceTitle: string;
  modified: Date | undefined;
}

export interface ICoverageResult {
  profile: ICoverageProfile;
  checkedAt: Date;
  sourceCount: number;
  searchCount: number;
  searchCountTrimmed: number;
  searchCountUntrimmed: number;
  duplicateDelta: number;
  delta: number;
  deltaPercent: number;
  sourceResults: ICoverageSourceResult[];
  missingSamples: ICoverageMissingItem[];
  warnings: string[];
  executedQueryText: string;
  executedQueryTemplate: string;
}

export interface IDiscoveredCoverageProfile {
  profile: ICoverageProfile;
  warnings: string[];
  sourcePageUrl: string;
}

interface ICoverageProfileInput {
  id?: string;
  title?: string;
  description?: string;
  queryTemplate?: string;
  resultSourceId?: string;
  sourceUrls?: string | string[];
  contentTypeIds?: string | string[];
  excludePaths?: string | string[];
  includeFolders?: boolean;
  trimDuplicates?: boolean;
  refinementFilters?: string;
}

interface IListInfoRaw {
  Title?: string;
  ItemCount?: number;
  NoCrawl?: boolean;
  Hidden?: boolean;
  RootFolder?: {
    ServerRelativeUrl?: string;
  };
}

interface ISourceItemRaw {
  Id: number;
  Title?: string;
  FileRef?: string;
  Modified?: string;
  ContentTypeId?: string;
  FileSystemObjectType?: number;
  FSObjType?: number;
}

interface IResolvedSourceItem {
  title: string;
  path: string;
  modified: Date | undefined;
  contentTypeId: string;
}

interface IResolvedSourceResult {
  title: string;
  sourceUrl: string;
  serverRelativeUrl: string;
  noCrawl: boolean;
  hidden: boolean;
  sourceCount: number;
  items: IResolvedSourceItem[];
}

interface IDiscoveredSearchResultsProperties {
  searchContextId?: string;
  queryTemplate?: string;
  resultSourceId?: string;
  refinementFilters?: string;
  refinementFiltersCollection?: Array<{
    property?: string;
    operator?: string;
    value?: string;
  }>;
  trimDuplicates?: boolean;
  searchScope?: string;
  searchScopePath?: string;
}

interface ICoverageCountSummary {
  effective: number;
  trimmed: number;
  untrimmed: number;
  duplicateDelta: number;
}

function createAbortError(): Error {
  const error = new Error('Aborted');
  error.name = 'AbortError';
  return error;
}

function uniqueValues(values: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];

  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    const normalized = value.toLowerCase();
    if (seen.has(normalized)) {
      continue;
    }
    seen.add(normalized);
    result.push(value);
  }

  return result;
}

function deriveWebUrlFromPageUrl(pageUrl: string): string | undefined {
  try {
    const parsed = new URL(pageUrl, window.location.origin);
    const marker = '/SitePages/';
    const markerIndex = parsed.pathname.toLowerCase().lastIndexOf(marker.toLowerCase());
    if (markerIndex >= 0) {
      return parsed.origin + parsed.pathname.substring(0, markerIndex);
    }
    return parsed.origin + parsed.pathname.replace(/\/[^/]+$/, '');
  } catch {
    return undefined;
  }
}

function deriveSiteCollectionUrlFromPageUrl(pageUrl: string): string | undefined {
  try {
    const parsed = new URL(pageUrl, window.location.origin);
    const segments = parsed.pathname.split('/').filter(Boolean);
    if (segments.length === 0) {
      return parsed.origin;
    }
    if ((segments[0] === 'sites' || segments[0] === 'teams') && segments.length >= 2) {
      return parsed.origin + '/' + segments[0] + '/' + segments[1];
    }
    return parsed.origin;
  } catch {
    return undefined;
  }
}

function toServerRelativePagePath(pageUrl: string): string {
  try {
    const parsed = new URL(pageUrl, window.location.origin);
    return parsed.pathname;
  } catch {
    return pageUrl;
  }
}

function buildRefinementFiltersString(
  refinementFiltersCollection: IDiscoveredSearchResultsProperties['refinementFiltersCollection']
): string | undefined {
  if (!refinementFiltersCollection || refinementFiltersCollection.length === 0) {
    return undefined;
  }

  const filters = refinementFiltersCollection
    .map(function (entry): string {
      const property = (entry.property || '').trim();
      const operator = (entry.operator || '').trim();
      const value = (entry.value || '').trim();
      if (!property || !operator || !value) {
        return '';
      }
      return property + ':' + operator + '(' + value + ')';
    })
    .filter(Boolean);

  return filters.length > 0 ? filters.join(',') : undefined;
}

export function buildCoverageProfileFromSearchResultsConfig(
  pageUrl: string,
  properties: IDiscoveredSearchResultsProperties,
  profileTitle?: string
): IDiscoveredCoverageProfile {
  const warnings: string[] = [];
  const sourceUrls: string[] = [];
  const searchScope = properties.searchScope || 'all';

  switch (searchScope) {
    case 'currentsite': {
      const webUrl = deriveWebUrlFromPageUrl(pageUrl);
      if (webUrl) {
        sourceUrls.push(webUrl);
        warnings.push('Source scope was auto-derived as the target page web. Replace it with specific libraries if this search experience is narrower.');
      }
      break;
    }
    case 'currentcollection': {
      const siteUrl = deriveSiteCollectionUrlFromPageUrl(pageUrl);
      if (siteUrl) {
        sourceUrls.push(siteUrl);
        warnings.push('Site collection scope was approximated from the target page URL. Add explicit libraries or site paths if the search experience is narrower.');
      }
      break;
    }
    case 'custom': {
      const customPath = (properties.searchScopePath || '').trim();
      if (customPath) {
        sourceUrls.push(customPath);
        warnings.push('Custom scope path was auto-detected from Search Results. Add explicit libraries if you need library-level coverage comparisons.');
      }
      break;
    }
    default:
      warnings.push('Search scope is set to All SharePoint, so source coverage cannot be inferred precisely from the query alone. Add explicit source paths.');
      break;
  }

  const refinementFilters = (properties.refinementFilters || '').trim() ||
    buildRefinementFiltersString(properties.refinementFiltersCollection);
  const normalizedPageUrl = normalizeAbsoluteUrl(pageUrl);

  return {
    profile: normalizeCoverageProfile({
      id: 'autodetect-' + normalizedPageUrl.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
      title: profileTitle || 'Auto-detected coverage profile',
      description: 'Derived from Search Results on ' + normalizedPageUrl,
      queryTemplate: properties.queryTemplate || DEFAULT_QUERY_TEMPLATE,
      resultSourceId: properties.resultSourceId,
      sourceUrls,
      trimDuplicates: properties.trimDuplicates !== false,
      refinementFilters
    }),
    warnings,
    sourcePageUrl: normalizedPageUrl
  };
}

export function normalizeDelimitedValues(value: string | string[] | undefined): string[] {
  if (Array.isArray(value)) {
    return uniqueValues(value.map(function (entry): string {
      return entry.trim();
    }).filter(Boolean));
  }

  if (!value) {
    return [];
  }

  return uniqueValues(value
    .split(/[\n,;]+/)
    .map(function (entry): string {
      return entry.trim();
    })
    .filter(Boolean));
}

export function normalizeCoverageProfile(profile: ICoverageProfileInput): ICoverageProfile {
  const title = (profile.title || '').trim();

  return {
    id: (profile.id || title || 'coverage-profile').trim(),
    title: title || 'Coverage Profile',
    description: (profile.description || '').trim() || undefined,
    queryTemplate: (profile.queryTemplate || DEFAULT_QUERY_TEMPLATE).trim() || DEFAULT_QUERY_TEMPLATE,
    resultSourceId: (profile.resultSourceId || '').trim() || undefined,
    sourceUrls: uniqueValues(normalizeDelimitedValues(profile.sourceUrls).map(function (sourceUrl): string {
      return normalizeAbsoluteUrl(sourceUrl);
    })),
    contentTypeIds: uniqueValues(normalizeDelimitedValues(profile.contentTypeIds)),
    excludePaths: uniqueValues(normalizeDelimitedValues(profile.excludePaths).map(function (path): string {
      return normalizeAbsoluteUrl(path);
    })),
    includeFolders: !!profile.includeFolders,
    trimDuplicates: profile.trimDuplicates !== false,
    refinementFilters: (profile.refinementFilters || '').trim() || undefined
  };
}

function normalizeAbsoluteUrl(value: string): string {
  if (!value) {
    return '';
  }

  try {
    const url = new URL(value, window.location.origin);
    return url.href.replace(/\/$/, '');
  } catch {
    return value.replace(/\/$/, '');
  }
}

function toServerRelativeUrl(value: string): string {
  try {
    const url = new URL(value, window.location.origin);
    return url.pathname.replace(/\/$/, '');
  } catch {
    return value.replace(window.location.origin, '').replace(/\/$/, '');
  }
}

function escapeKqlValue(value: string): string {
  return value.replace(/"/g, '\\"');
}

function isFolderItem(item: Pick<ISourceItemRaw, 'FileSystemObjectType' | 'FSObjType'>): boolean {
  if (typeof item.FileSystemObjectType === 'number') {
    return item.FileSystemObjectType === 1;
  }

  return item.FSObjType === 1;
}

export function matchesCoverageItem(
  profile: ICoverageProfile,
  item: {
    path?: string;
    contentTypeId?: string;
    isFolder?: boolean;
  }
): boolean {
  const path = normalizeAbsoluteUrl(item.path || '');

  if (!path) {
    return false;
  }

  if (!profile.includeFolders && item.isFolder) {
    return false;
  }

  if (profile.contentTypeIds.length > 0) {
    const itemContentTypeId = (item.contentTypeId || '').toLowerCase();
    const matchesContentType = profile.contentTypeIds.some(function (contentTypeId): boolean {
      return itemContentTypeId.indexOf(contentTypeId.toLowerCase()) === 0;
    });

    if (!matchesContentType) {
      return false;
    }
  }

  if (profile.excludePaths.length > 0) {
    const isExcluded = profile.excludePaths.some(function (excludePath): boolean {
      return path.toLowerCase().indexOf(excludePath.toLowerCase()) === 0;
    });

    if (isExcluded) {
      return false;
    }
  }

  return true;
}

export function buildCoverageQueryText(profile: ICoverageProfile, sourceUrls?: string[]): string {
  const clauses: string[] = ['*'];
  const effectiveSourceUrls = sourceUrls && sourceUrls.length > 0 ? sourceUrls : profile.sourceUrls;

  if (effectiveSourceUrls.length > 0) {
    clauses.push('(' + effectiveSourceUrls.map(function (sourceUrl): string {
      return 'Path:"' + escapeKqlValue(normalizeAbsoluteUrl(sourceUrl)) + '"';
    }).join(' OR ') + ')');
  }

  if (profile.contentTypeIds.length > 0) {
    clauses.push('(' + profile.contentTypeIds.map(function (contentTypeId): string {
      return 'ContentTypeId:' + escapeKqlValue(contentTypeId) + '*';
    }).join(' OR ') + ')');
  }

  if (profile.excludePaths.length > 0) {
    for (let i = 0; i < profile.excludePaths.length; i++) {
      clauses.push('NOT Path:"' + escapeKqlValue(profile.excludePaths[i]) + '"');
    }
  }

  return clauses.join(' AND ');
}

function buildItemValidationQueryText(profile: ICoverageProfile, itemPath: string): string {
  const clauses: string[] = ['Path:"' + escapeKqlValue(normalizeAbsoluteUrl(itemPath)) + '"'];

  if (profile.contentTypeIds.length > 0) {
    clauses.push('(' + profile.contentTypeIds.map(function (contentTypeId): string {
      return 'ContentTypeId:' + escapeKqlValue(contentTypeId) + '*';
    }).join(' OR ') + ')');
  }

  return clauses.join(' AND ');
}

export class SearchCoverageService {
  private readonly _provider: SharePointSearchProvider;

  public constructor() {
    this._provider = new SharePointSearchProvider();
  }

  public async discoverCoverageProfileFromPage(
    pageUrl: string,
    searchContextId?: string
  ): Promise<IDiscoveredCoverageProfile | undefined> {
    const serverRelativePagePath = toServerRelativePagePath(pageUrl);
    const page = await SPContext.sp.web.loadClientsidePage(serverRelativePagePath);
    const targetContextId = (searchContextId || 'default').trim() || 'default';

    for (let i = 0; i < page.sections.length; i++) {
      const section = page.sections[i];
      for (let j = 0; j < section.columns.length; j++) {
        const column = section.columns[j];
        for (let k = 0; k < column.controls.length; k++) {
          const control = column.controls[k] as unknown as {
            data?: {
              controlType?: number;
              webPartId?: string;
              webPartData?: {
                properties?: IDiscoveredSearchResultsProperties;
              };
            };
          };
          const webPartData = control.data;
          if (!webPartData || webPartData.controlType !== 3 || webPartData.webPartId !== SEARCH_RESULTS_WEBPART_ID) {
            continue;
          }

          const properties = webPartData.webPartData?.properties;
          if (!properties) {
            continue;
          }

          const controlContextId = (properties.searchContextId || 'default').trim() || 'default';
          if (controlContextId !== targetContextId) {
            continue;
          }

          return buildCoverageProfileFromSearchResultsConfig(
            pageUrl,
            properties,
            'Auto-detected from search page'
          );
        }
      }
    }

    return undefined;
  }

  public async evaluateProfile(profileInput: ICoverageProfile, signal: AbortSignal): Promise<ICoverageResult> {
    const profile = normalizeCoverageProfile(profileInput);
    const warningSet = new Set<string>();

    if (profile.sourceUrls.length === 0) {
      throw new Error('Add at least one list or library URL to the coverage profile.');
    }

    if (profile.queryTemplate !== DEFAULT_QUERY_TEMPLATE) {
      warningSet.add('Search counts apply the query template. Source counts only apply source URLs, content types, excluded paths, and the folder toggle.');
    }

    if (profile.refinementFilters) {
      warningSet.add('Persistent refinement filters are applied to search counts only.');
    }

    if (profile.trimDuplicates) {
      warningSet.add('Duplicate trimming can lower the search-side count compared to the source-side count.');
    }

    const resolvedSources: IResolvedSourceResult[] = [];
    const allSourceItems: Array<IResolvedSourceItem & { sourceTitle: string }> = [];

    for (let i = 0; i < profile.sourceUrls.length; i++) {
      if (signal.aborted) {
        throw createAbortError();
      }

      const configuredSourceUrl = profile.sourceUrls[i];

      try {
        const source = await this._loadSource(configuredSourceUrl, profile, signal);
        resolvedSources.push(source);

        if (source.noCrawl) {
          warningSet.add(source.title + ' is marked NoCrawl, so its items may be excluded from search until that setting changes.');
        }

        for (let j = 0; j < source.items.length; j++) {
          allSourceItems.push({
            ...source.items[j],
            sourceTitle: source.title
          });
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Failed to inspect source';
        warningSet.add('Unable to inspect ' + configuredSourceUrl + ': ' + message);
      }
    }

    if (resolvedSources.length === 0) {
      throw new Error('None of the configured coverage sources could be inspected.');
    }

    const normalizedSourceUrls = resolvedSources.map(function (source): string {
      return source.sourceUrl;
    });
    const overallSearchCount = await this._executeCountSummary(profile, normalizedSourceUrls, signal);

    const sourceResults: ICoverageSourceResult[] = [];
    let sourceCount = 0;

    for (let i = 0; i < resolvedSources.length; i++) {
      if (signal.aborted) {
        throw createAbortError();
      }

      const source = resolvedSources[i];
      const sourceSearchCount = await this._executeCountSummary(profile, [source.sourceUrl], signal);
      sourceCount += source.sourceCount;

      sourceResults.push({
        title: source.title,
        sourceUrl: source.sourceUrl,
        sourceCount: source.sourceCount,
        searchCount: sourceSearchCount.effective,
        searchCountTrimmed: sourceSearchCount.trimmed,
        searchCountUntrimmed: sourceSearchCount.untrimmed,
        duplicateDelta: sourceSearchCount.duplicateDelta,
        delta: source.sourceCount - sourceSearchCount.effective,
        noCrawl: source.noCrawl,
        hidden: source.hidden
      });
    }

    const missingSamples = await this._findMissingSamples(profile, allSourceItems, signal);
    const delta = sourceCount - overallSearchCount.effective;

    if (overallSearchCount.duplicateDelta > 0) {
      warningSet.add(
        'Duplicate collapsing changes the indexed count by ' + String(overallSearchCount.duplicateDelta) + ' item' +
        (overallSearchCount.duplicateDelta === 1 ? '' : 's') + '.'
      );
    }

    return {
      profile,
      checkedAt: new Date(),
      sourceCount,
      searchCount: overallSearchCount.effective,
      searchCountTrimmed: overallSearchCount.trimmed,
      searchCountUntrimmed: overallSearchCount.untrimmed,
      duplicateDelta: overallSearchCount.duplicateDelta,
      delta,
      deltaPercent: sourceCount > 0 ? Math.round((delta / sourceCount) * 1000) / 10 : 0,
      sourceResults,
      missingSamples,
      warnings: Array.from(warningSet),
      executedQueryText: buildCoverageQueryText(profile, normalizedSourceUrls),
      executedQueryTemplate: profile.queryTemplate
    };
  }

  private async _loadSource(
    sourceUrl: string,
    profile: ICoverageProfile,
    signal: AbortSignal
  ): Promise<IResolvedSourceResult> {
    try {
      return await this._loadListSource(sourceUrl, profile, signal);
    } catch {
      return this._loadWebSource(sourceUrl, profile, signal);
    }
  }

  private async _loadListSource(
    sourceUrl: string,
    profile: ICoverageProfile,
    signal: AbortSignal
  ): Promise<IResolvedSourceResult> {
    const listUrl = toServerRelativeUrl(sourceUrl);
    const list = SPContext.sp.web.getList(listUrl);
    const listInfo = await list
      .using(CacheNever())
      .select('Title', 'ItemCount', 'NoCrawl', 'Hidden', 'RootFolder/ServerRelativeUrl')
      .expand('RootFolder')<IListInfoRaw>();

    if (signal.aborted) {
      throw createAbortError();
    }

    const itemQuery = list.items
      .using(CacheNever())
      .select('Id', 'Title', 'FileRef', 'Modified', 'ContentTypeId', 'FileSystemObjectType', 'FSObjType')
      .top(5000);

    let page = await itemQuery.getPaged<ISourceItemRaw[]>();
    const items: ISourceItemRaw[] = [];

    while (true) {
      for (let i = 0; i < page.results.length; i++) {
        items.push(page.results[i]);
      }

      if (!page.hasNext) {
        break;
      }

      if (signal.aborted) {
        throw createAbortError();
      }

      const nextPage = await page.getNext();
      if (!nextPage) {
        break;
      }
      page = nextPage;
    }

    const sourceTitle = listInfo.Title || sourceUrl;
    const normalizedSourceUrl = normalizeAbsoluteUrl(
      listInfo.RootFolder?.ServerRelativeUrl || sourceUrl
    );
    const resolvedItems: IResolvedSourceItem[] = [];

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const path = normalizeAbsoluteUrl(item.FileRef || '');
      const modified = item.Modified ? new Date(item.Modified) : undefined;
      const includeItem = matchesCoverageItem(profile, {
        path,
        contentTypeId: item.ContentTypeId,
        isFolder: isFolderItem(item)
      });

      if (!includeItem) {
        continue;
      }

      resolvedItems.push({
        title: item.Title || path.split('/').pop() || 'Untitled',
        path,
        modified: modified && !isNaN(modified.getTime()) ? modified : undefined,
        contentTypeId: item.ContentTypeId || ''
      });
    }

    return {
      title: sourceTitle,
      sourceUrl: normalizedSourceUrl,
      serverRelativeUrl: listInfo.RootFolder?.ServerRelativeUrl || listUrl,
      noCrawl: !!listInfo.NoCrawl,
      hidden: !!listInfo.Hidden,
      sourceCount: resolvedItems.length,
      items: resolvedItems
    };
  }

  private async _loadWebSource(
    sourceUrl: string,
    profile: ICoverageProfile,
    signal: AbortSignal
  ): Promise<IResolvedSourceResult> {
    const normalizedSourceUrl = normalizeAbsoluteUrl(sourceUrl);
    const sourceRelativePath = toServerRelativeUrl(sourceUrl);
    const lists = await SPContext.sp.web.lists
      .using(CacheNever())
      .select('Title', 'Hidden', 'NoCrawl', 'BaseType', 'RootFolder/ServerRelativeUrl')
      .expand('RootFolder')();

    const sourceResults: IResolvedSourceItem[] = [];
    const sourceTitle = normalizedSourceUrl;

    for (let i = 0; i < lists.length; i++) {
      if (signal.aborted) {
        throw createAbortError();
      }

      const listInfo = lists[i] as IListInfoRaw & { BaseType?: number };
      const rootFolderUrl = listInfo.RootFolder?.ServerRelativeUrl || '';
      if (!rootFolderUrl || listInfo.Hidden || listInfo.BaseType === undefined) {
        continue;
      }
      if (listInfo.BaseType !== 0 && listInfo.BaseType !== 1) {
        continue;
      }
      if (rootFolderUrl.toLowerCase().indexOf(sourceRelativePath.toLowerCase()) !== 0) {
        continue;
      }

      const listResult = await this._loadListSource(rootFolderUrl, profile, signal);
      sourceResults.push(...listResult.items);
    }

    return {
      title: sourceTitle,
      sourceUrl: normalizedSourceUrl,
      serverRelativeUrl: sourceRelativePath,
      noCrawl: false,
      hidden: false,
      sourceCount: sourceResults.length,
      items: sourceResults
    };
  }

  private async _executeCountQuery(
    profile: ICoverageProfile,
    sourceUrls: string[],
    trimDuplicates: boolean,
    signal: AbortSignal
  ): Promise<number> {
    const query: ISearchQuery = {
      queryText: buildCoverageQueryText(profile, sourceUrls),
      queryTemplate: profile.queryTemplate,
      scope: COVERAGE_SCOPE,
      filters: [],
      sort: undefined,
      page: 1,
      pageSize: 1,
      selectedProperties: ['Path'],
      refiners: [],
      resultSourceId: profile.resultSourceId,
      trimDuplicates,
      enableQueryRules: false,
      refinementFilters: profile.refinementFilters
    };

    const response = await this._provider.execute(query, signal);
    return response.totalCount;
  }

  private async _executeCountSummary(
    profile: ICoverageProfile,
    sourceUrls: string[],
    signal: AbortSignal
  ): Promise<ICoverageCountSummary> {
    const trimmed = await this._executeCountQuery(profile, sourceUrls, true, signal);
    const untrimmed = await this._executeCountQuery(profile, sourceUrls, false, signal);

    return {
      effective: profile.trimDuplicates ? trimmed : untrimmed,
      trimmed,
      untrimmed,
      duplicateDelta: Math.max(0, untrimmed - trimmed)
    };
  }

  private async _findMissingSamples(
    profile: ICoverageProfile,
    items: Array<IResolvedSourceItem & { sourceTitle: string }>,
    signal: AbortSignal
  ): Promise<ICoverageMissingItem[]> {
    const sortedItems = items
      .slice()
      .sort(function (a, b): number {
        const aTime = a.modified ? a.modified.getTime() : 0;
        const bTime = b.modified ? b.modified.getTime() : 0;
        return bTime - aTime;
      });

    const missingItems: ICoverageMissingItem[] = [];

    for (let i = 0; i < sortedItems.length && i < MAX_MISSING_SAMPLE_SCAN; i++) {
      if (signal.aborted) {
        throw createAbortError();
      }

      const item = sortedItems[i];
      const query: ISearchQuery = {
        queryText: buildItemValidationQueryText(profile, item.path),
        queryTemplate: profile.queryTemplate,
        scope: COVERAGE_SCOPE,
        filters: [],
        sort: undefined,
        page: 1,
        pageSize: 1,
        selectedProperties: ['Path'],
        refiners: [],
        resultSourceId: profile.resultSourceId,
        trimDuplicates: profile.trimDuplicates,
        enableQueryRules: false,
        refinementFilters: profile.refinementFilters
      };
      const response = await this._provider.execute(query, signal);

      if (response.totalCount === 0) {
        missingItems.push({
          title: item.title,
          path: item.path,
          sourceTitle: item.sourceTitle,
          modified: item.modified
        });
      }

      if (missingItems.length >= MAX_MISSING_SAMPLE_ITEMS) {
        break;
      }
    }

    return missingItems;
  }
}
