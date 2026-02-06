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
 * Global store registry — maps searchContextId to context instances.
 * Web parts sharing the same searchContextId share the same store.
 */
const contextMap: Map<string, ISearchContext> = new Map();

/**
 * In-flight initialization promises — prevents race conditions when
 * multiple web parts call initializeSearchContext concurrently.
 */
const initPromises: Map<string, Promise<void>> = new Map();

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
  const context = contextMap.get(searchContextId);
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
 */
export async function initializeSearchContext(
  searchContextId: string
): Promise<void> {
  const context = getOrCreateContext(searchContextId);

  // Skip if already initialized
  if (context.isInitialized) {
    return;
  }

  // If initialization is already in-flight, await the existing promise
  const existing = initPromises.get(searchContextId);
  if (existing) {
    return existing;
  }

  // Create and start the initialization promise
  const promise = _doInitializeContext(searchContextId, context);
  initPromises.set(searchContextId, promise);

  try {
    await promise;
  } finally {
    initPromises.delete(searchContextId);
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
}

/**
 * Get or create a context for the given search context ID.
 */
function getOrCreateContext(searchContextId: string): ISearchContext {
  let context = contextMap.get(searchContextId);
  if (!context) {
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
    contextMap.set(searchContextId, context);
  }
  return context;
}

/**
 * Dispose and remove a store for the given search context ID.
 * Call this when all web parts using this context are unmounted.
 */
export function disposeStore(searchContextId: string): void {
  const context = contextMap.get(searchContextId);
  if (context) {
    // Tear down URL sync subscription (removes popstate listener + debounce timers)
    if (context.urlSyncUnsubscribe) {
      context.urlSyncUnsubscribe();
      context.urlSyncUnsubscribe = undefined;
    }
    context.orchestrator.stop();
    context.store.getState().dispose();
    contextMap.delete(searchContextId);
  }
}

/**
 * Check if a store exists for the given search context ID.
 */
export function hasStore(searchContextId: string): boolean {
  return contextMap.has(searchContextId);
}
