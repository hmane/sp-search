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

/**
 * Context instance that holds the store, orchestrator, and services.
 */
interface ISearchContext {
  store: StoreApi<ISearchStore>;
  orchestrator: SearchOrchestrator;
  managerService: SearchManagerService | undefined;
  urlSyncUnsubscribe: (() => void) | undefined;
  isInitialized: boolean;
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
  spfxContext?: any
): Promise<void> {
  const context = getOrCreateContext(searchContextId);

  // Initialize SPContext in the library's own webpack bundle.
  // Each web part bundle has its own copy of SPContext (due to webpack entry-point
  // duplication). The library's providers/services import SPContext from the library
  // bundle, which is a SEPARATE instance from the web part bundles. Without this
  // call, the library's SPContext stays uninitialized and throws at runtime.
  // SPContext.basic() is idempotent — second call returns existing context.
  if (spfxContext && !SPContext.isReady()) {
    await SPContext.basic(spfxContext, 'SpSearchStore');
  }

  // Skip if already initialized
  if (context.isInitialized) {
    console.log('[SP Search] v1.0.5 — initializeSearchContext("' + searchContextId + '") SKIPPED (already initialized)');
    return;
  }

  const promises = getInitPromises();

  // If initialization is already in-flight, await the existing promise
  const existing = promises.get(searchContextId);
  if (existing) {
    console.log('[SP Search] v1.0.5 — initializeSearchContext("' + searchContextId + '") AWAITING in-flight promise');
    return existing;
  }

  console.log('[SP Search] v1.0.5 — initializeSearchContext("' + searchContextId + '") STARTING initialization');

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

  // Create the URL sync subscription (namespace = searchContextId if not 'default')
  const urlPrefix = searchContextId !== 'default' ? searchContextId : undefined;
  context.urlSyncUnsubscribe = createUrlSyncSubscription(
    context.store as StoreApi<ISearchStore>,
    urlPrefix
  );

  // Resolve Azure AD group memberships for audience targeting (non-blocking)
  resolveUserGroupIds()
    .then(function (groupIds: string[]): void {
      context.store.getState().setCurrentUserGroups(groupIds);
    })
    .catch(function noop(): void { /* swallow — empty groups = fail-closed */ });

  context.isInitialized = true;

  // Trigger initial search — the orchestrator only reacts to state changes,
  // so on first page load (no URL params) we need to kick off the first search.
  // Uses queryText || '*' internally to match all items.
  context.orchestrator.triggerSearch().catch(function noop(): void { /* handled in orchestrator */ });
}

/**
 * Get or create a context for the given search context ID.
 */
function getOrCreateContext(searchContextId: string): ISearchContext {
  const map = getContextMap();
  let context = map.get(searchContextId);
  if (!context) {
    console.log('[SP Search] v1.0.5 — Creating NEW context for "' + searchContextId + '" (map size: ' + map.size + ')');
    const registries = createRegistryContainer();
    const store = createSearchStore(registries);
    const orchestrator = new SearchOrchestrator(store);
    context = {
      store,
      orchestrator,
      managerService: undefined,
      urlSyncUnsubscribe: undefined,
      isInitialized: false,
    };
    map.set(searchContextId, context);
  } else {
    console.log('[SP Search] v1.0.5 — Reusing EXISTING context for "' + searchContextId + '" (map size: ' + map.size + ')');
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
