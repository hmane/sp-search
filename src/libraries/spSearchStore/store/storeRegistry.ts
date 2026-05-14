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
import { initializeFileTypeIcons } from '@fluentui/react-file-type-icons';

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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const _win = window as any;

function getContextMap(): Map<string, ISearchContext> {
  if (!_win[CONTEXT_MAP_KEY]) {
    _win[CONTEXT_MAP_KEY] = new Map();
  }
  return _win[CONTEXT_MAP_KEY];
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
  }

  // Register Fluent UI file type icons (SVGs from Office CDN). Idempotent.
  initializeFileTypeIcons();

  // Skip if already initialized
  if (context.isInitialized) {
    console.log('[SP Search] v1.0.12 — initializeSearchContext("' + searchContextId + '") SKIPPED (already initialized)');
    return;
  }

  const promises = getInitPromises();

  // If initialization is already in-flight, await the existing promise
  const existing = promises.get(searchContextId);
  if (existing) {
    console.log('[SP Search] v1.0.12 — initializeSearchContext("' + searchContextId + '") AWAITING in-flight promise');
    return existing;
  }

  console.log('[SP Search] v1.0.12 — initializeSearchContext("' + searchContextId + '") STARTING initialization');

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
    console.log('[SP Search] v1.0.12 — Creating NEW context for "' + searchContextId + '" (map size: ' + map.size + ')');
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
    console.log('[SP Search] v1.0.12 — Reusing EXISTING context for "' + searchContextId + '" (map size: ' + map.size + ')');
  }
  return context;
}

/**
 * Dispose and remove a store for the given search context ID.
 * Call this when all web parts using this context are unmounted.
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
