import { ISearchScope, ISuggestion } from './ISearchScope';
import { IActiveFilter, IFilterConfig, IRefiner } from './IFilterTypes';
import { ISearchResult, IPersonaInfo, ISortField } from './ISearchResult';
import { IVerticalDefinition } from './IVerticalDefinition';

// ─── Query Slice ─────────────────────────────────────────────

export interface IQuerySlice {
  queryText: string;
  queryTemplate: string;
  /**
   * Client-side transformation applied to the user's raw input before execution.
   * Use {searchTerms} as the placeholder for the typed text.
   * Examples: "{searchTerms}*" (wildcard), "Title:{searchTerms}" (field-restricted).
   * Default: "{searchTerms}" (pass through unchanged).
   */
  queryInputTransformation: string;
  scope: ISearchScope;
  suggestions: ISuggestion[];
  isSearching: boolean;
  abortController: AbortController | undefined;
  // Actions
  setQueryText: (text: string) => void;
  setScope: (scope: ISearchScope) => void;
  setSuggestions: (suggestions: ISuggestion[]) => void;
  setQueryInputTransformation: (transformation: string) => void;
  cancelSearch: () => void;
}

// ─── Filter Slice ────────────────────────────────────────────

export interface IFilterSlice {
  activeFilters: IActiveFilter[];
  availableRefiners: IRefiner[];
  /** Refiner stability mode: debounced version for UI rendering */
  displayRefiners: IRefiner[];
  filterConfig: IFilterConfig[];
  isRefining: boolean;
  /**
   * How multiple selected filter groups (different properties) are combined.
   * 'AND' (default): results must match ALL selected filters.
   * 'OR': results matching ANY selected filter are returned.
   */
  operatorBetweenFilters: 'AND' | 'OR';
  // Actions
  setRefiner: (filter: IActiveFilter) => void;
  removeRefiner: (filterKey: string, value?: string) => void;
  clearAllFilters: () => void;
  setAvailableRefiners: (refiners: IRefiner[]) => void;
  setOperatorBetweenFilters: (operator: 'AND' | 'OR') => void;
}

// ─── Result Slice ────────────────────────────────────────────

export interface IPromotedResultItem {
  title: string;
  url: string;
  description?: string;
  iconUrl?: string;
}

export interface ISortableProperty {
  property: string;
  label: string;
  direction: string;
}

export interface IResultSlice {
  items: ISearchResult[];
  totalCount: number;
  currentPage: number;
  pageSize: number;
  sort: ISortField | undefined;
  /** Admin-configured sort options for the sort dropdown */
  sortableProperties: ISortableProperty[];
  promotedResults: IPromotedResultItem[];
  /** "Did you mean..." suggestion from the search API (QueryModification / SpellingSuggestion) */
  querySuggestion: string | undefined;
  isLoading: boolean;
  /**
   * True after the first search has completed (success or error).
   * Prevents showing "No results found" during the initial load before
   * any search has executed. Set to true by setResults and setError.
   */
  hasSearched: boolean;
  error: string | undefined;
  /** Search configuration — synced from web part property pane */
  resultSourceId: string;
  enableQueryRules: boolean;
  trimDuplicates: boolean;
  /** Persistent FQL refinement filters (always applied, separate from user filters) */
  refinementFilters: string;
  collapseSpecification: string;
  showPaging: boolean;
  pageRange: number;
  /** Additional managed properties to retrieve (merged with defaults) */
  selectedProperties: string;
  // Actions
  setResults: (items: ISearchResult[], total: number) => void;
  setPage: (page: number) => void;
  setSort: (sort: ISortField) => void;
  setPromotedResults: (results: IPromotedResultItem[]) => void;
  setQuerySuggestion: (suggestion: string | undefined) => void;
  setLoading: (isLoading: boolean) => void;
  setError: (error: string | undefined) => void;
}

// ─── Vertical Slice ──────────────────────────────────────────

export interface IVerticalSlice {
  currentVerticalKey: string;
  verticals: IVerticalDefinition[];
  verticalCounts: Record<string, number>;
  // Actions
  setVertical: (key: string) => void;
  setVerticalCounts: (counts: Record<string, number>) => void;
}

// ─── UI Slice ────────────────────────────────────────────────

export interface IUISlice {
  activeLayoutKey: string;
  /**
   * Layout keys available in the toolbar switcher.
   * Admins configure this to hide layouts not relevant to the search scenario.
   * Default: all six layouts.
   */
  availableLayouts: string[];
  isSearchManagerOpen: boolean;
  previewPanel: {
    isOpen: boolean;
    item: ISearchResult | undefined;
  };
  bulkSelection: string[];
  /** Current user's Azure AD security group IDs (for audience targeting) */
  currentUserGroups: string[];
  // Actions
  setLayout: (key: string) => void;
  setAvailableLayouts: (layouts: string[]) => void;
  toggleSearchManager: (isOpen?: boolean) => void;
  setPreviewItem: (item: ISearchResult | undefined) => void;
  toggleSelection: (itemKey: string, multiSelect: boolean) => void;
  clearSelection: () => void;
  setCurrentUserGroups: (groups: string[]) => void;
}

// ─── User Slice ──────────────────────────────────────────────

export interface ISavedSearch {
  id: number;
  title: string;
  queryText: string;
  /** JSON: serialized Query+Filter+Vertical slices */
  searchState: string;
  searchUrl: string;
  entryType: 'SavedSearch' | 'SharedSearch';
  category: string;
  sharedWith: IPersonaInfo[];
  resultCount: number;
  lastUsed: Date;
  created: Date;
  author: IPersonaInfo;
}

export interface ICollectionItem {
  /** SharePoint list item ID - required for unpin operations */
  id: number;
  url: string;
  title: string;
  metadata: Record<string, unknown>;
  sortOrder: number;
  /** User-defined tags/annotations for this item */
  tags: string[];
}

export interface ISearchCollection {
  id: number;
  collectionName: string;
  items: ICollectionItem[];
  sharedWith: IPersonaInfo[];
  tags: string[];
  created: Date;
  author: IPersonaInfo;
}

export interface IClickedItem {
  url: string;
  title: string;
  position: number;
  timestamp: Date;
}

export interface ISearchHistoryEntry {
  id: number;
  queryHash: string;
  queryText: string;
  vertical: string;
  scope: string;
  /** JSON-serialized full search state for restore */
  searchState: string;
  resultCount: number;
  isZeroResult?: boolean;
  clickedItems: IClickedItem[];
  searchTimestamp: Date;
}

export interface IUserSlice {
  savedSearches: ISavedSearch[];
  searchHistory: ISearchHistoryEntry[];
  collections: ISearchCollection[];
  // State update actions (called after service operations)
  setSavedSearches: (searches: ISavedSearch[]) => void;
  setSearchHistory: (history: ISearchHistoryEntry[]) => void;
  setCollections: (collections: ISearchCollection[]) => void;
  addSavedSearch: (search: ISavedSearch) => void;
  removeSavedSearch: (id: number) => void;
  addToHistory: (entry: ISearchHistoryEntry) => void;
  clearSearchHistory: () => void;
}
