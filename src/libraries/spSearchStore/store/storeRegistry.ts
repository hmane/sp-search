import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '@interfaces/index';
import { createRegistryContainer } from '@registries/index';
import { createSearchStore } from './createStore';
import { SearchOrchestrator } from '@orchestrator/SearchOrchestrator';
import { SearchManagerService, resolveUserGroupIds } from '@services/index';
import {
  createUrlSyncSubscription,
  setStateSnapshotHandler,
  setStateSnapshotLoader
} from './middleware';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { configureLegacyPnPBaseUrl } from 'spfx-toolkit/lib/utilities/context/urlSanitizer';
import { initializeFileTypeIcons } from '@fluentui/react-file-type-icons';
import { spLog } from '@store/utils/spLog';
import { seedToggleDefaults } from '../utils/toggleDefaults';

// Re-exported for back-compat with existing import paths (`@store/store/storeRegistry`).
// The implementation lives in ../utils/toggleDefaults to avoid an import cycle
// (this module imports the orchestrator + middleware, which now need the helper).
export { seedToggleDefaults };

/**
 * Context instance that holds the store, orchestrator, and services.
 */
interface ISearchContext {
  store: StoreApi<ISearchStore>;
  orchestrator: SearchOrchestrator;
  managerService: SearchManagerService | undefined;
  urlSyncUnsubscribe: (() => void) | undefined;
  isInitialized: boolean;
  /** Stable URL prefix computed at context creation time (before initialization). */
  urlPrefix: string;
  /**
   * T3.D6 — admin-provided URL prefix override. When set, wins over
   * the auto-computed `urlPrefix`. Admins configure short readable
   * prefixes like `ctx1` instead of the auto-generated 6+6-char hash.
   */
  urlPrefixOverride?: string;
  /**
   * T3.D6 — URL sync opt-out. When `false`, the context never
   * subscribes to URL changes / pushes state to the URL — useful
   * for embedded "saved-search runner" widgets that must not stomp
   * the page URL. Default true.
   */
  enableUrlSync: boolean;
}

// T3.D6 — initialization options surface admins can set per-context.
export interface IInitializeContextOptions {
  /** Explicit URL prefix override (e.g. "ctx1"). */
  urlPrefix?: string;
  /** Opt out of URL sync entirely. Default true. */
  enableUrlSync?: boolean;
}

/**
 * Window-backed global singletons for the store registry.
 *
 * SPFx bundles each web part as a separate webpack entry. Webpack aliases
 * (e.g. @store/*) cause this module to be duplicated into every web part
 * bundle, each with its own module-scoped variables. By storing the Maps
 * on `window`, ALL web part bundles share the same context instances.
 */
const CONTEXT_MAP_KEY = '__sp_search_context_map__';
const INIT_PROMISES_KEY = '__sp_search_init_promises__';
// T3.D1 — per-context refcount stored on window for cross-bundle visibility.
// Each web part bundle imports this module independently; without the
// window-backed map the refcounts would be per-bundle and never reach 0.
const REFCOUNT_KEY = '__sp_search_context_refcount_v1__';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const _win = window as any;

function getContextMap(): Map<string, ISearchContext> {
  if (!_win[CONTEXT_MAP_KEY]) {
    _win[CONTEXT_MAP_KEY] = new Map();
  }
  return _win[CONTEXT_MAP_KEY];
}

function getRefCountMap(): Map<string, number> {
  if (!_win[REFCOUNT_KEY]) {
    _win[REFCOUNT_KEY] = new Map<string, number>();
  }
  return _win[REFCOUNT_KEY];
}

function getInitPromises(): Map<string, Promise<void>> {
  if (!_win[INIT_PROMISES_KEY]) {
    _win[INIT_PROMISES_KEY] = new Map();
  }
  return _win[INIT_PROMISES_KEY];
}

/**
 * Get or create a store for the given search context ID.
 * First call creates the store; subsequent calls return the existing instance.
 */
export function getStore(searchContextId: string): StoreApi<ISearchStore> {
  return getOrCreateContext(searchContextId).store;
}

/**
 * Get or create the orchestrator for the given search context ID.
 */
export function getOrchestrator(searchContextId: string): SearchOrchestrator {
  return getOrCreateContext(searchContextId).orchestrator;
}

/**
 * Get the SearchManagerService for the given search context ID.
 * Returns undefined if initializeSearchContext has not been called.
 */
export function getManagerService(searchContextId: string): SearchManagerService | undefined {
  const context = getContextMap().get(searchContextId);
  return context?.managerService;
}

/**
 * Initialize the search context for full functionality.
 * Call this once from a web part's onInit() to enable:
 * - Search history logging
 * - Click tracking
 * - Saved searches, collections, etc.
 *
 * Uses a promise-based lock to prevent race conditions when
 * multiple web parts sharing the same context call this concurrently.
 *
 * @param searchContextId - The shared context ID
 * @param spfxContext - SPFx web part context (needed to initialize SPContext in the library bundle)
 */
export async function initializeSearchContext(
  searchContextId: string,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  spfxContext?: any,
  options?: IInitializeContextOptions
): Promise<void> {
  const context = getOrCreateContext(searchContextId);

  // T3.D6 — admin URL-prefix override + sync opt-out. These are
  // recorded on the context BEFORE the urlSync subscription is created
  // below. Re-init with different options is allowed; the most recent
  // wins because re-init creates a new subscription with the new prefix.
  if (options) {
    if (typeof options.urlPrefix === 'string') {
      const trimmed = options.urlPrefix.trim();
      context.urlPrefixOverride = trimmed.length > 0 ? trimmed : undefined;
    }
    if (typeof options.enableUrlSync === 'boolean') {
      context.enableUrlSync = options.enableUrlSync;
    }
  }

  // Initialize SPContext in the library's own webpack bundle.
  // Each web part bundle has its own copy of SPContext (due to webpack entry-point
  // duplication). The library's providers/services import SPContext from the library
  // bundle, which is a SEPARATE instance from the web part bundles. Without this
  // call, the library's SPContext stays uninitialized and throws at runtime.
  // SPContext.basic() is idempotent — second call returns existing context.
  if (spfxContext && !SPContext.isReady()) {
    await SPContext.basic(spfxContext, 'SpSearchStore');
    // Strip _layouts/15 contamination from the PnP v2 base URL bundled with
    // @pnp/spfx-controls-react + patch global fetch. Idempotent; safe to call
    // even when a web part's onInit already invoked it (separate webpack bundle).
    configureLegacyPnPBaseUrl(spfxContext);
  }

  // Register Fluent UI file type icons (SVGs from Office CDN). Idempotent.
  initializeFileTypeIcons();

  // Skip if already initialized
  if (context.isInitialized) {
    spLog.debug('initializeSearchContext skipped; context already initialized', { searchContextId });
    return;
  }

  const promises = getInitPromises();

  // If initialization is already in-flight, await the existing promise
  const existing = promises.get(searchContextId);
  if (existing) {
    spLog.debug('initializeSearchContext awaiting in-flight promise', { searchContextId });
    return existing;
  }

  spLog.debug('initializeSearchContext starting initialization', { searchContextId });

  // Create and start the initialization promise
  const promise = _doInitializeContext(searchContextId, context);
  promises.set(searchContextId, promise);

  try {
    await promise;
  } finally {
    promises.delete(searchContextId);
  }
}

/**
 * Internal initialization logic — separated to support promise-based locking.
 */
async function _doInitializeContext(
  searchContextId: string,
  context: ISearchContext
): Promise<void> {
  // Proactively clean up stale SPFx numeric hash keys from localStorage to prevent
  // QuotaExceededError. SPFx serializes web part property bags using numeric hash
  // keys that accumulate over time and can exceed the 5MB localStorage limit.
  _cleanupStaleStorage();

  // Create and initialize the SearchManagerService (uses SPContext.sp internally)
  const managerService = new SearchManagerService();
  await managerService.initialize();
  context.managerService = managerService;

  // Wire the history service to the orchestrator
  context.orchestrator.setHistoryService(managerService);

  // Start the orchestrator (listens to store changes)
  context.orchestrator.start();

  // Wire up URL sync — bi-directional state <-> URL
  // Set snapshot handlers for ?sid= fallback (long URLs)
  setStateSnapshotHandler(function (stateJson: string): Promise<number> {
    return managerService.saveStateSnapshot(stateJson);
  });
  setStateSnapshotLoader(function (stateId: number): Promise<string> {
    return managerService.loadStateSnapshot(stateId);
  });

  // T3.D6 — when the admin opts out of URL sync (`enableUrlSync: false`),
  // skip the subscription entirely. The store still functions; only URL
  // round-trips are disabled (useful for embedded saved-search runners
  // that must not stomp the page URL).
  if (context.enableUrlSync !== false) {
    // Create the URL sync subscription.
    // Prefer clean unprefixed URLs for the common single-search-page case.
    // If multiple independent search contexts exist, use the stored prefix.
    // T3.D6 — admin override wins over the auto-computed prefix.
    const autoPrefix = getContextMap().size > 1 ? context.urlPrefix : undefined;
    const effectivePrefix = context.urlPrefixOverride !== undefined ? context.urlPrefixOverride : autoPrefix;
    context.urlSyncUnsubscribe = createUrlSyncSubscription(
      context.store,
      effectivePrefix
    );
  }

  // Toggle defaultValue seeding — runs AFTER URL hydration so URL state
  // always wins over admin defaults (preserves shareable links). Pure
  // helper is exported separately for unit testing.
  //
  // The Filters web part syncs `filterConfig` before calling
  // initializeSearchContext, so configs are usually present here. But
  // when the Box/Results web part triggers init FIRST (before Filters
  // mounts), filterConfig is empty at this point — install a one-shot
  // subscription that runs the seed once filterConfig arrives.
  _runToggleDefaultSeed(context.store);

  // Resolve Azure AD group memberships for audience targeting (non-blocking)
  resolveUserGroupIds()
    .then(function (groupIds: string[]): void {
      context.store.getState().setCurrentUserGroups(groupIds);
    })
    .catch(function noop(): void { /* swallow — empty groups = fail-closed */ });

  context.isInitialized = true;

  // NOTE: We do NOT call orchestrator.triggerSearch() here.
  // The Results web part calls triggerSearch() at the end of its onInit(),
  // AFTER all providers/actions are registered and all configuration
  // (scope, queryTemplate, verticals, filterConfig) has been synced.
  // Calling it here would fire a premature search with incomplete state
  // (e.g. no vertical config, wrong scope) because the Results web part
  // often loads AFTER this initialization completes.
}

/**
 * Get or create a context for the given search context ID.
 */
function getOrCreateContext(searchContextId: string): ISearchContext {
  const map = getContextMap();
  let context = map.get(searchContextId);
  if (!context) {
    spLog.debug('Creating new search context', { searchContextId, contextCount: map.size });
    const registries = createRegistryContainer();
    const store = createSearchStore(registries);
    const orchestrator = new SearchOrchestrator(store);
    context = {
      store,
      orchestrator,
      managerService: undefined,
      urlSyncUnsubscribe: undefined,
      isInitialized: false,
      urlPrefix: _buildStableUrlPrefix(searchContextId),
      // T3.D6 — defaults; admins override via initializeSearchContext options.
      enableUrlSync: true,
      urlPrefixOverride: undefined,
    };
    map.set(searchContextId, context);

    // When a second context arrives (map.size transitions from 1 → 2+), the
    // first context was initialized without a URL prefix. Re-subscribe ALL
    // previously-initialized contexts so they pick up their stored prefix.
    // T3.D6 — also respect the admin override + enableUrlSync flag.
    if (map.size === 2) {
      map.forEach(function (existingCtx: ISearchContext, _existingId: string): void {
        if (existingCtx.isInitialized && existingCtx.urlSyncUnsubscribe) {
          existingCtx.urlSyncUnsubscribe();
          existingCtx.urlSyncUnsubscribe = undefined;
        }
        if (existingCtx.isInitialized && existingCtx.enableUrlSync !== false) {
          const prefix = existingCtx.urlPrefixOverride !== undefined
            ? existingCtx.urlPrefixOverride
            : existingCtx.urlPrefix;
          existingCtx.urlSyncUnsubscribe = createUrlSyncSubscription(
            existingCtx.store,
            prefix
          );
        }
      });
    }
  } else {
    spLog.debug('Reusing existing search context', { searchContextId, contextCount: map.size });
  }
  return context;
}

/**
 * T3.D1 — increment the per-context refcount. Call from each web part's
 * `onInit()` after `getStore(searchContextId)` resolves. The refcount is
 * keyed by `searchContextId` (NOT by web part instance) so that two web
 * parts sharing one context release together when both unmount.
 */
export function incrementContextRef(searchContextId: string): void {
  const refs = getRefCountMap();
  refs.set(searchContextId, (refs.get(searchContextId) || 0) + 1);
}

/**
 * T3.D1 — decrement the per-context refcount + schedule a deferred
 * dispose. Call from each web part's `onDispose()` BEFORE
 * `ReactDom.unmountComponentAtNode`.
 *
 * The dispose is deferred to a microtask so SPFx Modern's "next page
 * onInit fires before old page onDispose" cross-page navigation order
 * doesn't drop the context prematurely — a new mount has a chance to
 * re-increment first. If after the microtask the refcount is still 0,
 * the context is disposed.
 */
export function decrementContextRef(searchContextId: string): void {
  const refs = getRefCountMap();
  const current = refs.get(searchContextId) || 0;
  if (current <= 0) {
    // Already at zero — defensive; spurious onDispose with no prior
    // onInit (e.g. early SPFx unmount during a failed init).
    return;
  }
  const next = current - 1;
  if (next > 0) {
    refs.set(searchContextId, next);
    return;
  }
  // Transition to zero — schedule the dispose for the next microtask.
  // If a new mount with the same context arrives before then,
  // `incrementContextRef` will set the count back to 1 and the deferred
  // check below will skip the dispose.
  refs.set(searchContextId, 0);
  const scheduler: (cb: () => void) => void =
    typeof queueMicrotask === 'function' ? queueMicrotask : ((cb): void => { setTimeout(cb, 0); });
  scheduler((): void => {
    if ((refs.get(searchContextId) || 0) === 0) {
      refs.delete(searchContextId);
      disposeStore(searchContextId);
    }
  });
}

/**
 * Dispose and remove a store for the given search context ID.
 *
 * T3.D1 — prefer `decrementContextRef` from web part `onDispose()` so the
 * refcount + deferred-dispose contract handles the cross-page navigation
 * race. Direct calls to `disposeStore` bypass the refcount and force
 * immediate teardown — useful for tests and admin-driven "Force dispose"
 * surfaces (e.g. DebugPanel Multi-Context tab). Library-component
 * consumers (third-party extensions per `extensibility-guide.md`) should
 * use the refcounted path.
 */
export function disposeStore(searchContextId: string): void {
  const map = getContextMap();
  const context = map.get(searchContextId);
  if (context) {
    // Tear down URL sync subscription (removes popstate listener + debounce timers)
    if (context.urlSyncUnsubscribe) {
      context.urlSyncUnsubscribe();
      context.urlSyncUnsubscribe = undefined;
    }
    context.orchestrator.stop();
    context.store.getState().dispose();
    map.delete(searchContextId);
  }
  // T3.D1 — also clear the refcount so a future fresh mount starts
  // from 0 rather than inheriting any stale count.
  const refs = getRefCountMap();
  refs.delete(searchContextId);
}

/**
 * T3.D1 — test/admin helper: read the current refcount for a context.
 * Returns 0 when the context isn't tracked.
 */
export function getContextRefCount(searchContextId: string): number {
  return getRefCountMap().get(searchContextId) || 0;
}

/**
 * Check if a store exists for the given search context ID.
 */
export function hasStore(searchContextId: string): boolean {
  return getContextMap().has(searchContextId);
}

/**
 * Clean up stale SPFx localStorage entries to prevent QuotaExceededError.
 * SPFx serializes web part property bags using numeric hash keys (e.g. '1517419502').
 * Over time these accumulate and can exceed the 5MB localStorage limit.
 * This removes all numeric-keyed entries, which SPFx will regenerate as needed.
 */
function _cleanupStaleStorage(): void {
  try {
    const keysToRemove: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && /^-?\d+$/.test(key)) {
        keysToRemove.push(key);
      }
    }
    if (keysToRemove.length > 0) {
      for (let i = 0; i < keysToRemove.length; i++) {
        localStorage.removeItem(keysToRemove[i]);
      }
    }
  } catch {
    // localStorage not available or access denied — ignore
  }
}

/**
 * Apply toggle defaultValue seeding to a store. If filterConfig is
 * already populated, runs immediately. Otherwise, installs a one-shot
 * subscription that waits for filterConfig to arrive (the Filters web
 * part may mount AFTER initializeSearchContext was triggered by Box or
 * Results), then seeds and unsubscribes.
 *
 * Idempotent at the seed level — `seedToggleDefaults` skips any
 * managed property already present in activeFilters, so URL state and
 * any user interactions in the brief gap continue to win.
 */
function _runToggleDefaultSeed(store: StoreApi<ISearchStore>): void {
  function trySeed(): boolean {
    const state = store.getState();
    const configs = state.filterConfig || [];
    if (configs.length === 0) {
      return false;
    }
    const seeded = seedToggleDefaults(state.activeFilters, configs);
    if (seeded !== state.activeFilters) {
      store.setState({ activeFilters: seeded });
    }
    return true;
  }
  if (trySeed()) {
    return;
  }
  // filterConfig not yet populated — wait for it. One-shot subscription.
  const unsub = store.subscribe((state, prev): void => {
    if (state.filterConfig !== prev.filterConfig && (state.filterConfig || []).length > 0) {
      if (trySeed()) {
        unsub();
      }
    }
  });
}

function _buildStableUrlPrefix(searchContextId: string): string {
  const raw = searchContextId || 's';
  const normalized = raw.replace(/[^a-z0-9]/gi, '').toLowerCase() || 's';
  const stem = normalized.substring(0, Math.min(6, normalized.length));
  return stem + _stableHash(raw).toString(36).substring(0, 6);
}

function _stableHash(value: string): number {
  let hash = 2166136261;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}
