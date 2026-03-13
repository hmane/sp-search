/**
 * Active filter selection — dispatched to filterSlice when
 * a user selects/deselects a refiner value.
 */
export interface IActiveFilter {
  /** Managed property name, e.g. "RefinableString00" */
  filterName: string;
  /** Raw refinement token value (used for query) */
  value: string;
  /** Human-readable display text (from IRefinerValue.name) */
  displayValue?: string;
  operator: 'AND' | 'OR';
}

/**
 * Refiner bucket returned by the Search API.
 */
export interface IRefiner {
  filterName: string;
  values: IRefinerValue[];
}

/**
 * Single refiner value with count and selection state.
 */
export interface IRefinerValue {
  /** Human-readable label (may need async resolution for taxonomy/user types) */
  name: string;
  /** Raw refinement token for the query */
  value: string;
  count: number;
  isSelected: boolean;
}

export type FilterType =
  | 'checkbox'
  | 'daterange'
  | 'text'
  | 'people'
  | 'taxonomy'
  | 'tagbox'
  | 'slider'
  | 'toggle';

export type FilterOperator = 'AND' | 'OR';

export type SortBy = 'count' | 'alphabetical' | 'custom';
export type SortDirection = 'asc' | 'desc';

/**
 * Filter configuration — defines a single filter group
 * in the Search Filters web part property pane.
 */
export interface IFilterConfig {
  id: string;
  displayName: string;
  managedProperty: string;
  /** Optional short alias used in URLs, e.g. ft, ct, au */
  urlAlias?: string;
  filterType: FilterType;
  operator: FilterOperator;
  maxValues: number;
  defaultExpanded: boolean;
  showCount: boolean;
  sortBy: SortBy;
  sortDirection: SortDirection;
  /** Whether the user can select multiple values (default: true) */
  multiValues: boolean;
  /** Optional: slider range minimum */
  rangeMin?: number;
  /** Optional: slider range maximum */
  rangeMax?: number;
  /** Optional: slider step */
  rangeStep?: number;
  /** Optional: numeric formatting hint for slider values */
  rangeFormat?: 'number' | 'bytes' | 'currency';
  /** Optional: currency code for numeric formatting */
  currency?: string;
  /** Optional: taxonomy term set ID to drive the tree */
  termSetId?: string;
  /** Optional: taxonomy selection includes children */
  includeChildren?: boolean;
  /** Optional: boolean true label */
  trueLabel?: string;
  /** Optional: boolean false label */
  falseLabel?: string;
}

/**
 * Converts between raw refinement tokens, human-readable display,
 * and URL-safe strings. Each filter type provides a formatter.
 */
export interface IFilterValueFormatter {
  /** Unique formatter ID, matches the filter type ID */
  id: string;
  /** Convert raw refiner token → human-readable display text for the pill bar */
  formatForDisplay: (rawValue: string, config: IFilterConfig) => string | Promise<string>;
  /** Convert user selection → FQL/KQL refinement token for the Search API */
  formatForQuery: (displayValue: unknown, config: IFilterConfig) => string;
  /** Convert raw refiner token → URL-safe string for deep linking */
  formatForUrl: (rawValue: string) => string;
  /** Restore from URL-safe string → raw refiner token */
  parseFromUrl: (urlValue: string) => string;
}
