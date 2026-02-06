import * as React from 'react';

/**
 * Layout definition — registered via LayoutRegistry.
 * Each layout is lazy-loaded for code splitting.
 *
 * Built-in: DataGrid, Card, List, Compact, People, DocumentGallery
 */
export interface ILayoutDefinition {
  id: string;
  displayName: string;
  /** Fluent UI icon name for the layout switcher */
  iconName: string;
  /** Lazy-loaded React component */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  component: React.LazyExoticComponent<React.ComponentType<any>>;
  supportsPaging: 'numbered' | 'infinite' | 'both';
  supportsBulkSelect: boolean;
  supportsVirtualization: boolean;
  defaultSortable: boolean;
}
