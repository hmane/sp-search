// Core data types
export type { ISearchResult, IPersonaInfo, ISortField } from './ISearchResult';
export type { ISearchScope, ISuggestion } from './ISearchScope';

// Filter & refiner types
export type {
  IActiveFilter,
  IRefiner,
  IRefinerValue,
  IFilterConfig,
  IFilterValueFormatter,
  FilterType,
  FilterOperator,
  SortBy,
  SortDirection
} from './IFilterTypes';

// Vertical definition
export type { IVerticalDefinition } from './IVerticalDefinition';

// Store slices
export type {
  IQuerySlice,
  IFilterSlice,
  IResultSlice,
  ISortableProperty,
  IVerticalSlice,
  IUISlice,
  IUserSlice,
  IPromotedResultItem,
  ISavedSearch,
  ISearchCollection,
  ICollectionItem,
  ISearchHistoryEntry,
  IClickedItem
} from './IStoreSlices';

// Data provider
export type {
  ISearchDataProvider,
  ISearchQuery,
  ISearchResponse,
  IManagedProperty
} from './ISearchDataProvider';

// Suggestion provider
export type { ISuggestionProvider, ISearchContext } from './ISuggestionProvider';

// Action provider
export type { IActionProvider } from './IActionProvider';

// Layout & filter extensibility
export type { ILayoutDefinition } from './ILayoutDefinition';
export type { IFilterTypeDefinition } from './IFilterTypeDefinition';

// Promoted results
export type { IPromotedResultRule, IPromotedResultDisplay } from './IPromotedResult';

// Registry
export type { IRegistry, IRegistryContainer } from './IRegistry';

// Root store
export type { ISearchStore } from './ISearchStore';
