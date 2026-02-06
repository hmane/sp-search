import { ISearchResult, ISortField } from './ISearchResult';
import { ISearchScope, ISuggestion } from './ISearchScope';
import { IActiveFilter, IRefiner } from './IFilterTypes';
import { IPromotedResultItem } from './IStoreSlices';

// ─── Normalized Query Input ──────────────────────────────────

/**
 * Normalized search query — constructed by SearchService,
 * consumed by ISearchDataProvider implementations.
 */
export interface ISearchQuery {
  queryText: string;
  queryTemplate: string;
  scope: ISearchScope;
  filters: IActiveFilter[];
  sort: ISortField | undefined;
  page: number;
  pageSize: number;
  selectedProperties: string[];
  /** Managed properties to request refiner data for */
  refiners: string[];
  /** CollapseSpecification value — validated for sortability before use */
  collapseSpecification?: string;
  /** SharePoint-specific: result source GUID (ignored by Graph provider) */
  resultSourceId?: string;
  trimDuplicates?: boolean;
}

// ─── Normalized Query Output ─────────────────────────────────

/**
 * Normalized search response — every provider maps raw API
 * results to this interface.
 */
export interface ISearchResponse {
  items: ISearchResult[];
  totalCount: number;
  refiners: IRefiner[];
  promotedResults: IPromotedResultItem[];
  /** "Did you mean..." suggestion from the search API */
  querySuggestion?: string;
}

// ─── Managed Property Schema ─────────────────────────────────

/**
 * Managed property metadata from the Search Administration API.
 * Used by the Schema Helper property pane control.
 */
export interface IManagedProperty {
  name: string;
  /** Text, DateTime, Integer, Double, etc. */
  type: string;
  /** Human-readable alias if configured */
  alias?: string;
  queryable: boolean;
  retrievable: boolean;
  refinable: boolean;
  sortable: boolean;
}

// ─── Data Provider Interface ─────────────────────────────────

/**
 * Abstraction over search backends. Web parts never call
 * PnPjs or Graph directly — they always go through a provider.
 *
 * Built-in: SharePointSearchProvider, GraphSearchProvider
 */
export interface ISearchDataProvider {
  id: string;
  displayName: string;
  /** Execute a search and return normalized results */
  execute: (query: ISearchQuery, signal: AbortSignal) => Promise<ISearchResponse>;
  /** Optional: provider-specific suggestions */
  getSuggestions?: (query: string, signal: AbortSignal) => Promise<ISuggestion[]>;
  /** Optional: fetch available properties for Schema Helper */
  getSchema?: () => Promise<IManagedProperty[]>;
  supportsRefiners: boolean;
  supportsCollapsing: boolean;
  supportsSorting: boolean;
}
