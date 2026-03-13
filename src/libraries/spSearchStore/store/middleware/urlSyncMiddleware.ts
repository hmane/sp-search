import { StoreApi } from 'zustand/vanilla';
import {
  ISearchStore,
  IActiveFilter,
  IFilterConfig,
  ISortField,
  ISearchScope,
} from '@interfaces/index';
import { getFilterValueFormatter } from '@store/formatters/FilterValueFormatters';
import { getFilterUrlAlias, sanitizeUrlAlias } from '@store/utils/filterUrlAliases';

// ─── URL State Shape ────────────────────────────────────────────

/**
 * Minimal state shape that is round-tripped through the URL.
 * Only the fields the user explicitly changed are serialized;
 * defaults (page 1, layout 'list') are omitted to keep URLs clean.
 */
interface IUrlState {
  queryText?: string;
  activeFilters?: IActiveFilter[];
  urlFilters?: IUrlFilterParam[];
  currentVerticalKey?: string;
  sort?: ISortField;
  currentPage?: number;
  scope?: string;
  activeLayoutKey?: string;
  /** If set, a ?sid= was found — caller should load the snapshot and apply it */
  stateId?: number;
}

interface IUrlFilterParam {
  key: string;
  rawValue: string;
}

// ─── URL Param Keys ─────────────────────────────────────────────

/** Short parameter names kept terse for readable URLs. */
const PARAM_QUERY = 'q';
const PARAM_FILTERS = 'f';
const PARAM_VERTICAL = 'v';
const PARAM_SORT = 's';
const PARAM_PAGE = 'p';
const PARAM_SCOPE = 'c';
const PARAM_LAYOUT = 'l';
const PARAM_STATE_VERSION = 'x';
const PARAM_STATE_ID = 'i';
const LEGACY_PARAM_SCOPE = 'sc';
const LEGACY_PARAM_STATE_VERSION = 'sv';
const LEGACY_PARAM_STATE_ID = 'sid';

/** Maximum URL length before falling back to ?sid= deep link. */
const MAX_URL_LENGTH = 2000;

/** Debounce delay (ms) for URL pushes to avoid excessive history entries. */
const DEBOUNCE_MS = 300;

// ─── Store Type ─────────────────────────────────────────────────

/**
 * Store shape consumed by the URL sync middleware.
 * URL filter serialization/deserialization depends on filterConfig
 * and registered filter type definitions, so this uses the full store.
 */
type IUrlSyncStoreSlice = ISearchStore;

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

function getParam(params: URLSearchParams, key: string, prefix?: string, legacyKey?: string): string | null {
  const currentValue = params.get(prefixKey(key, prefix));
  if (currentValue !== null) {
    return currentValue;
  }
  return legacyKey ? params.get(prefixKey(legacyKey, prefix)) : null;
}

const RESERVED_PARAM_KEYS = new Set([
  PARAM_QUERY,
  PARAM_FILTERS,
  PARAM_VERTICAL,
  PARAM_SORT,
  PARAM_PAGE,
  PARAM_SCOPE,
  PARAM_LAYOUT,
  PARAM_STATE_VERSION,
  PARAM_STATE_ID,
  LEGACY_PARAM_SCOPE,
  LEGACY_PARAM_STATE_VERSION,
  LEGACY_PARAM_STATE_ID,
]);

function getFilterParamKey(urlKey: string, prefix?: string): string {
  return prefix ? `${prefix}.${urlKey}` : urlKey;
}

function normalizeFilterKey(value: string | undefined): string {
  return sanitizeUrlAlias(value) || '';
}

function isReservedParam(key: string, prefix?: string): boolean {
  if (!prefix) {
    return RESERVED_PARAM_KEYS.has(key);
  }
  if (key.indexOf(prefix + '.') !== 0) {
    return false;
  }
  return RESERVED_PARAM_KEYS.has(key.substring(prefix.length + 1));
}

function clearFilterParams(
  params: URLSearchParams,
  state: Partial<IUrlSyncStoreSlice>,
  prefix?: string
): void {
  const keysToDelete: string[] = [];
  const filterConfig = state.filterConfig || [];
  const filterKeys = new Set<string>();

  for (let i = 0; i < filterConfig.length; i++) {
    filterKeys.add(getFilterParamKey(getFilterUrlAlias(filterConfig[i]), prefix));
    filterKeys.add(prefixKey(filterConfig[i].managedProperty, prefix));
  }

  if (state.activeFilters) {
    for (let i = 0; i < state.activeFilters.length; i++) {
      const filter = state.activeFilters[i];
      const config = filterConfig.find((f) => f.managedProperty === filter.filterName);
      filterKeys.add(getFilterParamKey(getFilterUrlAlias(config || {
        managedProperty: filter.filterName,
        filterType: 'checkbox',
      } as IFilterConfig), prefix));
      filterKeys.add(prefixKey(filter.filterName, prefix));
    }
  }

  params.forEach((_value, key) => {
    if (filterKeys.has(key) || key.indexOf(prefixKey('f.', prefix)) === 0) {
      keysToDelete.push(key);
    }
  });
  for (let i = 0; i < keysToDelete.length; i++) {
    params.delete(keysToDelete[i]);
  }
}

function extractPeopleEmail(value: string): string {
  const lower = value.toLowerCase();
  if (lower.indexOf('|membership|') >= 0) {
    return lower.substring(lower.lastIndexOf('|') + 1);
  }
  return value;
}

function decodeUrlComponentSafely(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function fromBase64(encoded: string): string {
  try {
    return decodeURIComponent(escape(atob(encoded)));
  } catch {
    return '';
  }
}

function stripWrappingQuotes(value: string): string {
  if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
    return value.substring(1, value.length - 1);
  }
  return value;
}

function decodeHexRefinementToken(value: string): string {
  const stripped = stripWrappingQuotes(value);
  if (stripped.indexOf('\u01C2\u01C2') !== 0) {
    return stripped;
  }
  const hex = stripped.substring(2);
  if (!/^[0-9a-fA-F]+$/.test(hex) || hex.length % 2 !== 0) {
    return stripped;
  }
  try {
    let encoded = '';
    for (let i = 0; i < hex.length; i += 2) {
      encoded += '%' + hex.substring(i, i + 2);
    }
    return decodeURIComponent(encoded);
  } catch {
    return stripped;
  }
}

function compactUrlFilterValue(value: string): string {
  return decodeHexRefinementToken(stripWrappingQuotes(value));
}

function parseLegacyFiltersParam(filtersRaw: string): IActiveFilter[] | undefined {
  try {
    const decoded = fromBase64(filtersRaw);
    if (!decoded) {
      return undefined;
    }

    const parsed: unknown = JSON.parse(decoded);
    if (!Array.isArray(parsed)) {
      return undefined;
    }

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
          displayValue: typeof item.displayValue === 'string' ? item.displayValue : undefined,
          operator: item.operator
        });
      }
    }

    return filters.length > 0 ? filters : undefined;
  } catch {
    return undefined;
  }
}

function serializeActiveFilterForUrl(
  filter: IActiveFilter,
  state: Partial<IUrlSyncStoreSlice>
): string {
  const filterConfig = state.filterConfig || [];
  let config: IFilterConfig | undefined;
  for (let i = 0; i < filterConfig.length; i++) {
    if (normalizeFilterKey(filterConfig[i].managedProperty) === normalizeFilterKey(filter.filterName)) {
      config = filterConfig[i];
      break;
    }
  }

  return compactUrlFilterValue(
    decodeUrlComponentSafely(
      getFilterValueFormatter(config?.filterType).formatForUrl(filter.value)
    )
  );
}

function shouldUseMultiValueParam(config: IFilterConfig | undefined): boolean {
  if (!config) {
    return false;
  }
  if (config.multiValues === false) {
    return false;
  }
  return config.filterType !== 'text' && config.filterType !== 'daterange' && config.filterType !== 'slider' && config.filterType !== 'toggle';
}

function splitSerializedFilterValues(rawValue: string, config: IFilterConfig | undefined): string[] {
  if (!rawValue) {
    return [];
  }
  if (!shouldUseMultiValueParam(config)) {
    return [rawValue];
  }
  return rawValue
    .split(',')
    .map((part) => decodeUrlComponentSafely(part.trim()))
    .filter((part) => part.length > 0);
}

function resolveFilterConfigByUrlKey(
  key: string,
  state: IUrlSyncStoreSlice
): IFilterConfig | undefined {
  const normalizedKey = normalizeFilterKey(key);
  const filterConfig = state.filterConfig || [];
  for (let i = 0; i < filterConfig.length; i++) {
    const config = filterConfig[i];
    if (
      normalizeFilterKey(getFilterUrlAlias(config)) === normalizedKey ||
      normalizeFilterKey(config.managedProperty) === normalizedKey
    ) {
      return config;
    }
  }
  return undefined;
}

function deserializeUrlFilterToActiveFilter(
  rawFilter: IUrlFilterParam,
  state: IUrlSyncStoreSlice,
  config?: IFilterConfig
): IActiveFilter | undefined {
  const resolvedConfig = config || resolveFilterConfigByUrlKey(rawFilter.key, state);

  if (!resolvedConfig) {
    return undefined;
  }

  const operator = resolvedConfig.operator || 'OR';

  const typeDef = state.registries.filterTypes.get(resolvedConfig.filterType);
  if (!typeDef) {
    return undefined;
  }

  const formatter = getFilterValueFormatter(resolvedConfig.filterType);
  const urlParsed = formatter.parseFromUrl(rawFilter.rawValue);
  const deserialized = typeDef.deserializeValue(urlParsed);
  let token = typeDef.buildRefinementToken(deserialized, resolvedConfig.managedProperty);
  let displayValue: string | undefined;

  if (resolvedConfig.filterType === 'people') {
    const raw = String(deserialized).trim();
    if (raw && raw.indexOf('|') < 0 && raw.indexOf('@') >= 0) {
      token = 'i:0#.f|membership|' + raw.toLowerCase();
      displayValue = raw;
    } else {
      displayValue = extractPeopleEmail(raw);
    }
  } else if (resolvedConfig.filterType === 'text') {
    displayValue = String(deserialized);
  }

  return {
    filterName: resolvedConfig.managedProperty,
    value: token,
    displayValue,
    operator
  };
}

function resolveUrlFilters(
  urlFilters: IUrlFilterParam[],
  state: IUrlSyncStoreSlice
): {
  resolved: IActiveFilter[];
  unresolved: IUrlFilterParam[];
} {
  const resolved: IActiveFilter[] = [];
  const unresolved: IUrlFilterParam[] = [];

  for (let i = 0; i < urlFilters.length; i++) {
    const rawFilter = urlFilters[i];
    const config = resolveFilterConfigByUrlKey(rawFilter.key, state);
    if (!config || !state.registries.filterTypes.get(config.filterType)) {
      unresolved.push(rawFilter);
      continue;
    }

    const parts = splitSerializedFilterValues(rawFilter.rawValue, config);
    for (let j = 0; j < parts.length; j++) {
      const next = deserializeUrlFilterToActiveFilter(
        {
          key: rawFilter.key,
          rawValue: parts[j],
        },
        state,
        config
      );
      if (next) {
        resolved.push(next);
      }
    }
  }

  return { resolved, unresolved };
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

  // ── q = queryText ──
  if (state.queryText) {
    params.set(prefixKey(PARAM_QUERY, prefix), state.queryText);
  } else {
    params.delete(prefixKey(PARAM_QUERY, prefix));
  }

  // ── <alias> = filter values (multi-select collapsed as comma-separated) ──
  clearFilterParams(params, state, prefix);
  if (state.activeFilters && state.activeFilters.length > 0) {
    const filterConfig = state.filterConfig || [];
    const groupedValues = new Map<string, string[]>();
    for (let i = 0; i < state.activeFilters.length; i++) {
      const filter = state.activeFilters[i];
      const config = filterConfig.find(
        (f) => normalizeFilterKey(f.managedProperty) === normalizeFilterKey(filter.filterName)
      );
      const key = getFilterParamKey(getFilterUrlAlias(config || {
        managedProperty: filter.filterName,
        filterType: 'checkbox',
      } as IFilterConfig), prefix);
      const serializedValue = serializeActiveFilterForUrl(filter, state);
      if (shouldUseMultiValueParam(config)) {
        const existing = groupedValues.get(key) || [];
        existing.push(serializedValue);
        groupedValues.set(key, existing);
      } else {
        params.set(key, serializedValue);
      }
    }

    groupedValues.forEach((values, key) => {
      params.set(key, values.map((value) => encodeURIComponent(value)).join(','));
    });
  }

  // If a non-grouped key was also grouped later, grouped value wins.
  if (state.activeFilters && state.activeFilters.length > 0) {
    const filterConfig = state.filterConfig || [];
    const groupedKeys = new Set<string>();
    for (let i = 0; i < state.activeFilters.length; i++) {
      const filter = state.activeFilters[i];
      const config = filterConfig.find(
        (f) => normalizeFilterKey(f.managedProperty) === normalizeFilterKey(filter.filterName)
      );
      if (shouldUseMultiValueParam(config)) {
        groupedKeys.add(getFilterParamKey(getFilterUrlAlias(config || {
          managedProperty: filter.filterName,
          filterType: 'checkbox',
        } as IFilterConfig), prefix));
      }
    }
    groupedKeys.forEach((key) => {
      const value = params.get(key);
      if (value !== null) {
        params.set(key, value);
      }
    });
  }

  // ── v = currentVerticalKey (only when not 'all') ──
  if (state.currentVerticalKey && state.currentVerticalKey !== 'all') {
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
  // Only serialize scope when it's a user-facing change (not the web part
  // property pane default). The scope is re-synced from property pane settings
  // on every page load, so persisting the default to the URL is unnecessary
  // and causes issues when the URL sync overwrites the scope before the
  // Results web part can set the full scope object (including kqlPath).
  // Scope is serialized only when it differs from common defaults.
  if (state.scope && state.scope.id !== 'all' && state.scope.id !== 'currentsite') {
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

  // Check for ?sid= (StateId fallback) first
  const stateIdRaw = getParam(params, PARAM_STATE_ID, prefix, LEGACY_PARAM_STATE_ID);
  if (stateIdRaw) {
    const stateId = parseInt(stateIdRaw, 10);
    if (!isNaN(stateId) && stateId > 0) {
      state.stateId = stateId;
      return state;
    }
  }

  // ── q ──
  const queryText = params.get(prefixKey(PARAM_QUERY, prefix));
  if (queryText) {
    state.queryText = queryText;
  }

  // ── legacy f = base64(JSON) activeFilters ──
  const legacyFilters = params.get(prefixKey(PARAM_FILTERS, prefix));
  if (legacyFilters) {
    state.activeFilters = parseLegacyFiltersParam(legacyFilters);
  }

  // ── <alias> / <managedProperty> filter params ──
  const urlFilters: IUrlFilterParam[] = [];
  params.forEach((value, key) => {
    if (isReservedParam(key, prefix)) {
      return;
    }
    let filterKey = key;
    if (prefix) {
      if (key.indexOf(prefix + '.') !== 0) {
        return;
      }
      filterKey = key.substring(prefix.length + 1);
    }
    if (!filterKey) {
      return;
    }
    urlFilters.push({
      key: filterKey,
      rawValue: value
    });
  });
  if (urlFilters.length > 0) {
    state.urlFilters = urlFilters;
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
  const scopeId = getParam(params, PARAM_SCOPE, prefix, LEGACY_PARAM_SCOPE);
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
 * Optional callback for saving state snapshots when URL exceeds MAX_URL_LENGTH.
 * Set via `setStateSnapshotHandler` before creating subscriptions.
 */
let _saveSnapshotHandler: ((stateJson: string) => Promise<number>) | undefined;

/**
 * Register the handler used to persist state snapshots for the ?sid= fallback.
 * Should be called during web part onInit with SearchManagerService.saveStateSnapshot.
 */
export function setStateSnapshotHandler(handler: (stateJson: string) => Promise<number>): void {
  _saveSnapshotHandler = handler;
}

/**
 * Push the current store state into the URL via `history.replaceState`.
 * The call is debounced so rapid successive state changes (e.g. typing)
 * collapse into a single URL update.
 *
 * If the resulting URL exceeds 2,000 characters and a snapshot handler
 * is registered, the full state is saved to the SearchSavedQueries list
 * and the URL is replaced with ?sid=<itemId>.
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

    // Check if URL exceeds max length — fall back to ?sid= if handler available
    if (newUrl.length > MAX_URL_LENGTH && _saveSnapshotHandler) {
      const stateJson = JSON.stringify({
        queryText: state.queryText,
        activeFilters: state.activeFilters,
        currentVerticalKey: state.currentVerticalKey,
        sort: state.sort,
        currentPage: state.currentPage,
        scope: state.scope,
        activeLayoutKey: state.activeLayoutKey,
      });

      _saveSnapshotHandler(stateJson)
        .then(function (stateId: number): void {
          if (stateId > 0) {
            const sidParams = new URLSearchParams();
            sidParams.set(prefixKey(PARAM_STATE_ID, prefix), String(stateId));
            const sidUrl = window.location.pathname + '?' + sidParams.toString() + window.location.hash;
            window.history.replaceState(window.history.state, '', sidUrl);
          }
        })
        .catch(function (): void {
          // Fallback: use the long URL anyway
          window.history.replaceState(window.history.state, '', newUrl);
        });
    } else {
      window.history.replaceState(
        window.history.state,
        '',
        newUrl
      );
    }

    delete debounceTimers[timerKey];
  }, DEBOUNCE_MS) as unknown as number;
}

// ─── Subscription Factory ───────────────────────────────────────

/**
 * Optional callback for loading state snapshots when ?sid= is detected.
 * Set via `setStateSnapshotLoader` before creating subscriptions.
 */
let _loadSnapshotHandler: ((stateId: number) => Promise<string>) | undefined;

/**
 * Register the handler used to load state snapshots for the ?sid= fallback.
 * Should be called during web part onInit with SearchManagerService.loadStateSnapshot.
 */
export function setStateSnapshotLoader(handler: (stateId: number) => Promise<string>): void {
  _loadSnapshotHandler = handler;
}

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
  let pendingUrlFilters: IUrlFilterParam[] | undefined;

  // Handle ?sid= fallback: load state snapshot from list
  if (initial.stateId && _loadSnapshotHandler) {
    _loadSnapshotHandler(initial.stateId)
      .then(function (stateJson: string): void {
        if (!stateJson) {
          return;
        }
        try {
          const parsed = JSON.parse(stateJson) as IUrlState;
          // sid= snapshots are not namespaced — no prefix needed here
          applyUrlStateToStore(store, parsed);
        } catch {
          // Malformed snapshot JSON — ignore
        }
      })
      .catch(function (): void {
        // Failed to load snapshot — stay with defaults
      });
  } else if (Object.keys(initial).length > 0 && !initial.stateId) {
    pendingUrlFilters = applyUrlStateToStore(store, initial, prefix).unresolvedUrlFilters;
  }

  // ─── 2. Store → URL subscription ───────────────────────────
  /**
   * Zustand subscribe returns an unsubscribe function.
   * We track a shallow snapshot of the URL-relevant fields to avoid
   * pushing when unrelated state changes (e.g. isLoading).
   */
  let previousSnapshot = takeSnapshot(store.getState());

  const unsubStore = store.subscribe((state): void => {
    if (pendingUrlFilters && pendingUrlFilters.length > 0) {
      if (state.activeFilters.length > 0) {
        pendingUrlFilters = undefined;
      } else if (state.filterConfig.length > 0) {
        const replayResult = applyUrlStateToStore(
          store,
          { urlFilters: pendingUrlFilters },
          prefix
        );
        pendingUrlFilters = replayResult.unresolvedUrlFilters;
        if (replayResult.didUpdate) {
          return;
        }
      }
    }

    const snapshot = takeSnapshot(state);
    if (!shallowEqualSnapshot(previousSnapshot, snapshot)) {
      previousSnapshot = snapshot;
      pushStateToUrl(state, prefix);
    }
  });

  // ─── 3. URL → Store (popstate) ─────────────────────────────
  const onPopState = (): void => {
    const urlState = deserializeFromUrl(prefix);
    pendingUrlFilters = applyUrlStateToStore(store, urlState, prefix).unresolvedUrlFilters;
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
 *
 * `availableLayouts` wins over `?l=`: if the URL requests a layout that
 * is not in the store's `availableLayouts`, `activeLayoutKey` is coerced
 * to the first available layout and the URL is normalized immediately so
 * the deep link reflects reality.
 *
 * @param prefix - Optional namespace prefix; passed through for URL normalization.
 */
function applyUrlStateToStore(
  store: StoreApi<IUrlSyncStoreSlice>,
  urlState: IUrlState,
  prefix?: string
): {
  didUpdate: boolean;
  unresolvedUrlFilters?: IUrlFilterParam[];
} {
  const patch: Record<string, unknown> = {};
  let unresolvedUrlFilters: IUrlFilterParam[] | undefined;

  if (urlState.queryText !== undefined) {
    patch.queryText = urlState.queryText;
  }

  if (urlState.urlFilters && urlState.urlFilters.length > 0) {
    const resolution = resolveUrlFilters(urlState.urlFilters, store.getState());
    if (resolution.resolved.length > 0) {
      patch.activeFilters = resolution.resolved;
    }
    if (resolution.unresolved.length > 0) {
      unresolvedUrlFilters = resolution.unresolved;
    }
  } else if (urlState.activeFilters !== undefined) {
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
    // availableLayouts wins — coerce if the URL requests a disabled layout.
    const availableLayouts: string[] = store.getState().availableLayouts;
    const requested = urlState.activeLayoutKey;
    const resolved = (availableLayouts.length > 0 && availableLayouts.indexOf(requested) >= 0)
      ? requested
      : (availableLayouts[0] || 'list');
    patch.activeLayoutKey = resolved;

    // Normalize the URL immediately when the requested layout was coerced,
    // so the deep link reflects the actual active layout rather than a
    // disabled layout key that would confuse copy-paste sharing.
    if (resolved !== requested && isBrowser()) {
      const lKey = prefixKey(PARAM_LAYOUT, prefix);
      const params = new URLSearchParams(window.location.search);
      // 'list' is the default — omit the param rather than writing ?l=list
      if (resolved !== 'list') {
        params.set(lKey, resolved);
      } else {
        params.delete(lKey);
      }
      const qs = params.toString();
      const normalized = qs
        ? `${window.location.pathname}?${qs}${window.location.hash}`
        : `${window.location.pathname}${window.location.hash}`;
      window.history.replaceState(window.history.state, '', normalized);
    }
  }

  if (Object.keys(patch).length > 0) {
    store.setState(patch as Partial<IUrlSyncStoreSlice>);
    return {
      didUpdate: true,
      unresolvedUrlFilters,
    };
  }

  return {
    didUpdate: false,
    unresolvedUrlFilters,
  };
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
