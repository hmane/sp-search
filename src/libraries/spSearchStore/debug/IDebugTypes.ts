export type DebugEventType = 'SEARCH' | 'FILTER' | 'VERTICAL' | 'URL' | 'ERROR' | 'INIT';

/**
 * T5.D2 — per-network-call entry rendered in the DebugPanel's Network
 * tab. The orchestrator's `_executeProviderWithRetry` chokepoint emits
 * one entry per provider call (main search + vertical count queries =
 * 4-6 per user-initiated search on a typical page).
 *
 * `queryText` is always stored as `[redacted]` — the panel surface is
 * admin-gated but the field carries literal user input and the audit's
 * acceptance signal requires it scrubbed before storage.
 */
export interface INetworkEvent {
  readonly id: number;
  readonly timestamp: number;
  readonly providerId: string;
  readonly kind: 'search' | 'verticalCount';
  readonly status: 'ok' | 'error' | 'aborted';
  /** Wall-clock duration in milliseconds, undefined if still in flight. */
  readonly durationMs: number | undefined;
  readonly queryText: string;
  readonly queryTemplate: string;
  readonly currentPage: number;
  readonly pageSize: number;
  readonly totalCount: number | undefined;
  readonly itemCount: number | undefined;
  readonly errorMessage: string | undefined;
  /** Optional sub-key used by the vertical-count fan-out (e.g. `documents`). */
  readonly verticalKey: string | undefined;
}

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
