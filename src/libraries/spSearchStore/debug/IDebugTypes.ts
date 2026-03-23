export type DebugEventType = 'SEARCH' | 'FILTER' | 'VERTICAL' | 'URL' | 'ERROR' | 'INIT';

export interface IDebugEvent {
  readonly id: number;
  readonly type: DebugEventType;
  readonly timestamp: number;
  readonly data: Record<string, unknown>;
}

export interface IQueryDebugInfo {
  readonly kql: string;
  readonly queryTemplate: string;
  readonly resultSourceId: string | undefined;
  readonly refinementFilters: string[];
  readonly providerId: string;
  readonly startTime: number;
  readonly duration: number | undefined;
  readonly totalCount: number | undefined;
  readonly itemsReturned: number | undefined;
  readonly currentPage: number;
  readonly pageSize: number;
  readonly refiners: Array<{ name: string; values: Array<{ value: string; count: number }> }>;
  readonly error: string | undefined;
  /** Full normalized request object (ISearchQuery) */
  readonly request: Record<string, unknown> | undefined;
  /** Full normalized response object (ISearchResponse summary) */
  readonly response: Record<string, unknown> | undefined;
}

export interface IWebPartDebugConfig {
  readonly componentName: string;
  readonly properties: Record<string, unknown>;
  readonly registeredAt: number;
}
