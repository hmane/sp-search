/**
 * T3.D10 — initialization-order diagnostic.
 *
 * Multi-context pages where Results loads before Filters cause the
 * first search to run with an empty `filterConfig`, so URL-deep-link
 * filter values silently fail to apply. The diagnostic records (a)
 * which web parts have called `initializeSearchContext`, (b) whether
 * the orchestrator has fired its first `triggerSearch`, (c) whether
 * `filterConfig` was empty at first-search time.
 *
 * The Results web part reads the diagnostic in edit mode and surfaces
 * a MessageBar + Retry button.
 *
 * Window-backed so all bundles see the same data.
 */

const REGISTRY_KEY = '__sp_search_init_order_diagnostic_v1__';

export interface IInitOrderDiagnostic {
  /** Web part class names that have called initializeSearchContext for this context. */
  registeredWebParts: Set<string>;
  /** True after the orchestrator fires its first triggerSearch. */
  firstSearchFired: boolean;
  /** filterConfig.length at first-search time. -1 means "not yet measured". */
  filterConfigAtFirstSearch: number;
  /** Whether the Filters web part registered AFTER first-search fired. */
  filtersLateRegistered: boolean;
}

interface IDiagnosticRegistry {
  /** Map<searchContextId, IInitOrderDiagnostic>. */
  byContext: Map<string, IInitOrderDiagnostic>;
}

interface IWindowWithRegistry {
  [REGISTRY_KEY]?: IDiagnosticRegistry;
}

function getRegistry(): IDiagnosticRegistry {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const win = (typeof window !== 'undefined' ? window : ({} as any)) as IWindowWithRegistry;
  let entry = win[REGISTRY_KEY];
  if (!entry) {
    entry = { byContext: new Map<string, IInitOrderDiagnostic>() };
    win[REGISTRY_KEY] = entry;
  }
  return entry;
}

function ensureContext(searchContextId: string): IInitOrderDiagnostic {
  const reg = getRegistry();
  let d = reg.byContext.get(searchContextId);
  if (!d) {
    d = {
      registeredWebParts: new Set<string>(),
      firstSearchFired: false,
      filterConfigAtFirstSearch: -1,
      filtersLateRegistered: false,
    };
    reg.byContext.set(searchContextId, d);
  }
  return d;
}

/**
 * Record that a web part has called `initializeSearchContext` for this
 * context. Filters web part calls with `webPartName='SpSearchFiltersWebPart'`;
 * the diagnostic raises `filtersLateRegistered` when this call arrives
 * AFTER `firstSearchFired === true`.
 */
export function recordWebPartInit(searchContextId: string, webPartName: string): void {
  const d = ensureContext(searchContextId);
  d.registeredWebParts.add(webPartName);
  if (webPartName === 'SpSearchFiltersWebPart' && d.firstSearchFired) {
    d.filtersLateRegistered = true;
  }
}

/**
 * Record that the orchestrator has fired its first `triggerSearch` for
 * this context. The Results web part calls this from its
 * `triggerSearch()` invocation site.
 */
export function recordFirstSearch(searchContextId: string, filterConfigLength: number): void {
  const d = ensureContext(searchContextId);
  if (!d.firstSearchFired) {
    d.firstSearchFired = true;
    d.filterConfigAtFirstSearch = filterConfigLength;
  }
}

/**
 * Read the diagnostic for a context. Returns `undefined` when no
 * diagnostic has been recorded yet (no web part has registered).
 */
export function getInitOrderDiagnostic(searchContextId: string): IInitOrderDiagnostic | undefined {
  return getRegistry().byContext.get(searchContextId);
}

/**
 * True when the diagnostic indicates a potential init-order issue:
 * Filters web part registered AFTER first search fired AND that first
 * search ran with empty filterConfig.
 */
export function hasInitOrderIssue(searchContextId: string): boolean {
  const d = getRegistry().byContext.get(searchContextId);
  if (!d) { return false; }
  return d.filtersLateRegistered && d.filterConfigAtFirstSearch === 0;
}

/**
 * Clear the diagnostic for a context. Used by the in-product Retry
 * button so a re-rendered search resets the state.
 */
export function clearInitOrderDiagnostic(searchContextId: string): void {
  getRegistry().byContext.delete(searchContextId);
}

/** Test-only — clears the entire registry. */
export function _resetInitOrderDiagnosticForTesting(): void {
  getRegistry().byContext.clear();
}
