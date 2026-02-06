import { StoreApi } from 'zustand/vanilla';
import {
  IActiveFilter,
  ISortField,
  ISearchScope,
  IQuerySlice,
  IFilterSlice,
  IResultSlice,
  IVerticalSlice,
  IUISlice
} from '@interfaces/index';

// ─── URL State Shape ────────────────────────────────────────────

/**
 * Minimal state shape that is round-tripped through the URL.
 * Only the fields the user explicitly changed are serialized;
 * defaults (page 1, layout 'list') are omitted to keep URLs clean.
 */
interface IUrlState {
  queryText?: string;
  activeFilters?: IActiveFilter[];
  currentVerticalKey?: string;
  sort?: ISortField;
  currentPage?: number;
  scope?: string;
  activeLayoutKey?: string;
}

// ─── URL Param Keys ─────────────────────────────────────────────

/** Short parameter names kept terse for readable URLs. */
const PARAM_QUERY = 'q';
const PARAM_FILTERS = 'f';
const PARAM_VERTICAL = 'v';
const PARAM_SORT = 's';
const PARAM_PAGE = 'p';
const PARAM_SCOPE = 'sc';
const PARAM_LAYOUT = 'l';
const PARAM_STATE_VERSION = 'sv';

/** Current state version — bumped if the schema changes. */
const STATE_VERSION = '1';

/** Debounce delay (ms) for URL pushes to avoid excessive history entries. */
const DEBOUNCE_MS = 300;

// ─── Store Type ─────────────────────────────────────────────────

/**
 * Minimal store shape consumed by the URL sync middleware.
 * We avoid importing the full ISearchStore to keep coupling low.
 */
type IUrlSyncStoreSlice =
  IQuerySlice &
  IFilterSlice &
  IResultSlice &
  IVerticalSlice &
  IUISlice;

// ─── Helpers ────────────────────────────────────────────────────

/**
 * Returns `true` when running in a browser context with `window` available.
 * Prevents crashes during SSR or Node-based test harnesses.
 */
function isBrowser(): boolean {
  return typeof window !== 'undefined' && typeof window.location !== 'undefined';
}

/**
 * Prefix a parameter name with the optional namespace.
 * When `prefix` is provided the key becomes `{prefix}.{key}`.
 */
function prefixKey(key: string, prefix?: string): string {
  return prefix ? `${prefix}.${key}` : key;
}

/**
 * Encode a value as base64.
 * Uses `btoa` which is available in all modern browsers.
 */
function toBase64(value: string): string {
  try {
    return btoa(unescape(encodeURIComponent(value)));
  } catch {
    return '';
  }
}

/**
 * Decode a base64 value.
 */
function fromBase64(encoded: string): string {
  try {
    return decodeURIComponent(escape(atob(encoded)));
  } catch {
    return '';
  }
}

// ─── Serialization ──────────────────────────────────────────────

/**
 * Serialize URL-relevant store state into the current `URLSearchParams`.
 * Existing params outside the middleware's namespace are preserved.
 *
 * @param state - The combined store state (or a partial snapshot).
 * @param prefix - Optional namespace prefix for multi-context pages.
 * @returns The full query-string (without leading `?`).
 */
export function serializeToUrl(
  state: Partial<IUrlSyncStoreSlice>,
  prefix?: string
): string {
  if (!isBrowser()) {
    return '';
  }

  const params = new URLSearchParams(window.location.search);

  // ── State version tag (always written) ──
  params.set(prefixKey(PARAM_STATE_VERSION, prefix), STATE_VERSION);

  // ── q = queryText ──
  if (state.queryText) {
    params.set(prefixKey(PARAM_QUERY, prefix), state.queryText);
  } else {
    params.delete(prefixKey(PARAM_QUERY, prefix));
  }

  // ── f = activeFilters (JSON → base64) ──
  if (state.activeFilters && state.activeFilters.length > 0) {
    const json = JSON.stringify(state.activeFilters);
    params.set(prefixKey(PARAM_FILTERS, prefix), toBase64(json));
  } else {
    params.delete(prefixKey(PARAM_FILTERS, prefix));
  }

  // ── v = currentVerticalKey ──
  if (state.currentVerticalKey) {
    params.set(prefixKey(PARAM_VERTICAL, prefix), state.currentVerticalKey);
  } else {
    params.delete(prefixKey(PARAM_VERTICAL, prefix));
  }

  // ── s = sort (property:direction) ──
  if (state.sort) {
    params.set(
      prefixKey(PARAM_SORT, prefix),
      `${state.sort.property}:${state.sort.direction}`
    );
  } else {
    params.delete(prefixKey(PARAM_SORT, prefix));
  }

  // ── p = currentPage (only when > 1) ──
  if (state.currentPage !== undefined && state.currentPage > 1) {
    params.set(prefixKey(PARAM_PAGE, prefix), String(state.currentPage));
  } else {
    params.delete(prefixKey(PARAM_PAGE, prefix));
  }

  // ── sc = scope id ──
  if (state.scope) {
    params.set(prefixKey(PARAM_SCOPE, prefix), state.scope.id);
  } else {
    params.delete(prefixKey(PARAM_SCOPE, prefix));
  }

  // ── l = activeLayoutKey (only when not 'list') ──
  if (state.activeLayoutKey && state.activeLayoutKey !== 'list') {
    params.set(prefixKey(PARAM_LAYOUT, prefix), state.activeLayoutKey);
  } else {
    params.delete(prefixKey(PARAM_LAYOUT, prefix));
  }

  return params.toString();
}

// ─── Deserialization ────────────────────────────────────────────

/**
 * Read URL parameters and return a partial state object that can
 * be applied to the store. Fields that are missing or invalid in
 * the URL are returned as `undefined` so the store keeps its defaults.
 *
 * @param prefix - Optional namespace prefix for multi-context pages.
 * @returns Partial URL state (only fields present in the URL are set).
 */
export function deserializeFromUrl(prefix?: string): IUrlState {
  if (!isBrowser()) {
    return {};
  }

  const params = new URLSearchParams(window.location.search);
  const state: IUrlState = {};

  // Bail early if no state-version tag — nothing was serialized by us
  // Note: URLSearchParams.get() returns null for missing keys.
  // We use truthiness checks which handle both null and empty string.
  const version = params.get(prefixKey(PARAM_STATE_VERSION, prefix));
  if (!version) {
    return state;
  }

  // ── q ──
  const queryText = params.get(prefixKey(PARAM_QUERY, prefix));
  if (queryText) {
    state.queryText = queryText;
  }

  // ── f ──
  const filtersRaw = params.get(prefixKey(PARAM_FILTERS, prefix));
  if (filtersRaw) {
    try {
      const decoded = fromBase64(filtersRaw);
      if (decoded) {
        const parsed: unknown = JSON.parse(decoded);
        if (Array.isArray(parsed)) {
          // Validate each item has the expected shape
          const filters: IActiveFilter[] = [];
          for (let i = 0; i < parsed.length; i++) {
            const item = parsed[i] as Record<string, unknown>;
            if (
              typeof item.filterName === 'string' &&
              typeof item.value === 'string' &&
              (item.operator === 'AND' || item.operator === 'OR')
            ) {
              filters.push({
                filterName: item.filterName,
                value: item.value,
                operator: item.operator
              });
            }
          }
          if (filters.length > 0) {
            state.activeFilters = filters;
          }
        }
      }
    } catch {
      // Malformed base64 or JSON — ignore silently
    }
  }

  // ── v ──
  const verticalKey = params.get(prefixKey(PARAM_VERTICAL, prefix));
  if (verticalKey) {
    state.currentVerticalKey = verticalKey;
  }

  // ── s ──
  const sortRaw = params.get(prefixKey(PARAM_SORT, prefix));
  if (sortRaw) {
    const colonIdx = sortRaw.lastIndexOf(':');
    if (colonIdx > 0) {
      const property = sortRaw.substring(0, colonIdx);
      const direction = sortRaw.substring(colonIdx + 1);
      if (direction === 'Ascending' || direction === 'Descending') {
        state.sort = { property, direction };
      }
    }
  }

  // ── p ──
  const pageRaw = params.get(prefixKey(PARAM_PAGE, prefix));
  if (pageRaw) {
    const page = parseInt(pageRaw, 10);
    if (!isNaN(page) && page >= 1) {
      state.currentPage = page;
    }
  }

  // ── sc ──
  const scopeId = params.get(prefixKey(PARAM_SCOPE, prefix));
  if (scopeId) {
    state.scope = scopeId;
  }

  // ── l ──
  const layoutKey = params.get(prefixKey(PARAM_LAYOUT, prefix));
  if (layoutKey) {
    state.activeLayoutKey = layoutKey;
  }

  return state;
}

// ─── URL Push (debounced replaceState) ──────────────────────────

/** Active debounce timer handle, keyed by prefix to support multiple contexts. */
const debounceTimers: Record<string, number> = {};

/**
 * Push the current store state into the URL via `history.replaceState`.
 * The call is debounced so rapid successive state changes (e.g. typing)
 * collapse into a single URL update.
 */
function pushStateToUrl(
  state: Partial<IUrlSyncStoreSlice>,
  prefix?: string
): void {
  if (!isBrowser()) {
    return;
  }

  const timerKey = prefix || '__default__';

  // Clear any pending debounce for this prefix
  if (debounceTimers[timerKey] !== undefined) {
    clearTimeout(debounceTimers[timerKey]);
  }

  debounceTimers[timerKey] = setTimeout((): void => {
    const qs = serializeToUrl(state, prefix);
    const newUrl = qs
      ? `${window.location.pathname}?${qs}${window.location.hash}`
      : `${window.location.pathname}${window.location.hash}`;

    window.history.replaceState(
      window.history.state,
      '',
      newUrl
    );

    delete debounceTimers[timerKey];
  }, DEBOUNCE_MS) as unknown as number;
}

// ─── Subscription Factory ───────────────────────────────────────

/**
 * Create a bi-directional subscription that keeps the URL and
 * the Zustand store in sync.
 *
 * **Store → URL** — subscribes to store changes and debounces URL writes.
 * **URL → Store** — listens for `popstate` (browser back/forward) and
 * rehydrates the store from the URL.
 *
 * @param store - The Zustand vanilla store instance.
 * @param prefix - Optional namespace prefix for multi-context pages.
 * @returns An `unsubscribe` function to tear down all listeners — call
 *          this inside `dispose()`.
 */
export function createUrlSyncSubscription(
  store: StoreApi<IUrlSyncStoreSlice>,
  prefix?: string
): () => void {
  if (!isBrowser()) {
    // SSR: return a no-op unsubscribe
    return (): void => { /* no-op */ };
  }

  // ─── 1. Hydrate store from URL on init ──────────────────────
  const initial = deserializeFromUrl(prefix);
  if (Object.keys(initial).length > 0) {
    applyUrlStateToStore(store, initial);
  }

  // ─── 2. Store → URL subscription ───────────────────────────
  /**
   * Zustand subscribe returns an unsubscribe function.
   * We track a shallow snapshot of the URL-relevant fields to avoid
   * pushing when unrelated state changes (e.g. isLoading).
   */
  let previousSnapshot = takeSnapshot(store.getState());

  const unsubStore = store.subscribe((state): void => {
    const snapshot = takeSnapshot(state);
    if (!shallowEqualSnapshot(previousSnapshot, snapshot)) {
      previousSnapshot = snapshot;
      pushStateToUrl(state, prefix);
    }
  });

  // ─── 3. URL → Store (popstate) ─────────────────────────────
  const onPopState = (): void => {
    const urlState = deserializeFromUrl(prefix);
    applyUrlStateToStore(store, urlState);
  };

  window.addEventListener('popstate', onPopState);

  // ─── 4. Return combined unsubscribe ────────────────────────
  return (): void => {
    unsubStore();
    window.removeEventListener('popstate', onPopState);

    // Clear any pending debounce timer
    const timerKey = prefix || '__default__';
    if (debounceTimers[timerKey] !== undefined) {
      clearTimeout(debounceTimers[timerKey]);
      delete debounceTimers[timerKey];
    }
  };
}

// ─── Internal Helpers ───────────────────────────────────────────

/**
 * Apply a deserialized URL state to the store, mapping IUrlState
 * fields to their corresponding slice properties.
 */
function applyUrlStateToStore(
  store: StoreApi<IUrlSyncStoreSlice>,
  urlState: IUrlState
): void {
  const patch: Record<string, unknown> = {};

  if (urlState.queryText !== undefined) {
    patch.queryText = urlState.queryText;
  }

  if (urlState.activeFilters !== undefined) {
    patch.activeFilters = urlState.activeFilters;
  }

  if (urlState.currentVerticalKey !== undefined) {
    patch.currentVerticalKey = urlState.currentVerticalKey;
  }

  if (urlState.sort !== undefined) {
    patch.sort = urlState.sort;
  }

  if (urlState.currentPage !== undefined) {
    patch.currentPage = urlState.currentPage;
  }

  if (urlState.scope !== undefined) {
    // The store expects a full ISearchScope. We can only restore the id
    // from the URL; the consuming code should resolve the label later.
    const currentScope = store.getState().scope;
    if (currentScope && currentScope.id === urlState.scope) {
      // Already matches — no need to update
    } else {
      patch.scope = { id: urlState.scope, label: urlState.scope } as ISearchScope;
    }
  }

  if (urlState.activeLayoutKey !== undefined) {
    patch.activeLayoutKey = urlState.activeLayoutKey;
  }

  if (Object.keys(patch).length > 0) {
    store.setState(patch as Partial<IUrlSyncStoreSlice>);
  }
}

/**
 * Snapshot of URL-relevant fields for shallow comparison.
 * Keeps the subscription from pushing to the URL on unrelated state changes.
 */
interface IUrlSnapshot {
  queryText: string;
  activeFilters: IActiveFilter[];
  currentVerticalKey: string;
  sort: ISortField | undefined;
  currentPage: number;
  scopeId: string;
  activeLayoutKey: string;
}

function takeSnapshot(state: IUrlSyncStoreSlice): IUrlSnapshot {
  return {
    queryText: state.queryText,
    activeFilters: state.activeFilters,
    currentVerticalKey: state.currentVerticalKey,
    sort: state.sort,
    currentPage: state.currentPage,
    scopeId: state.scope.id,
    activeLayoutKey: state.activeLayoutKey
  };
}

/**
 * Shallow comparison of two URL snapshots.
 * `activeFilters` is compared by reference (Zustand replaces the array
 * on every mutation, so reference equality is sufficient).
 */
function shallowEqualSnapshot(a: IUrlSnapshot, b: IUrlSnapshot): boolean {
  return (
    a.queryText === b.queryText &&
    a.activeFilters === b.activeFilters &&
    a.currentVerticalKey === b.currentVerticalKey &&
    a.sort === b.sort &&
    a.currentPage === b.currentPage &&
    a.scopeId === b.scopeId &&
    a.activeLayoutKey === b.activeLayoutKey
  );
}
