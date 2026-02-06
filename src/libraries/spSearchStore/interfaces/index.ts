// Core data types
export { ISearchResult, IPersonaInfo, ISortField } from './ISearchResult';
export { ISearchScope, ISuggestion } from './ISearchScope';

// Filter & refiner types
export {
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
export { IVerticalDefinition } from './IVerticalDefinition';

// Store slices
export {
  IQuerySlice,
  IFilterSlice,
  IResultSlice,
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
export {
  ISearchDataProvider,
  ISearchQuery,
  ISearchResponse,
  IManagedProperty
} from './ISearchDataProvider';

// Suggestion provider
export { ISuggestionProvider, ISearchContext } from './ISuggestionProvider';

// Action provider
export { IActionProvider } from './IActionProvider';

// Layout & filter extensibility
export { ILayoutDefinition } from './ILayoutDefinition';
export { IFilterTypeDefinition } from './IFilterTypeDefinition';

// Promoted results
export { IPromotedResultRule, IPromotedResultDisplay } from './IPromotedResult';

// Registry
export { IRegistry, IRegistryContainer } from './IRegistry';

// Root store
export { ISearchStore } from './ISearchStore';
