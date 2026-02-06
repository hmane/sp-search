import * as React from 'react';

/**
 * Filter type definition — registered via FilterTypeRegistry.
 * Each filter type is lazy-loaded for code splitting.
 *
 * Built-in: checkbox, daterange, people, taxonomy, tagbox, slider, toggle
 */
export interface IFilterTypeDefinition {
  id: string;
  displayName: string;
  /** Lazy-loaded filter UI component (createLazyComponent or React.lazy) */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  component: React.ComponentType<any>;
  /** Convert filter value to URL-safe string for deep linking */
  serializeValue: (value: unknown) => string;
  /** Restore filter value from URL param */
  deserializeValue: (raw: string) => unknown;
  /** Convert to KQL/FQL refinement token for the Search API */
  buildRefinementToken: (value: unknown, managedProperty: string) => string;
}
