import { StoreApi } from 'zustand/vanilla';
import { SPFI } from '@pnp/sp';
import { ISearchStore } from '@interfaces/index';
import { createRegistryContainer } from '@registries/index';
import { createSearchStore } from './createStore';
import { SearchOrchestrator } from '@orchestrator/SearchOrchestrator';
import { SearchManagerService } from '@services/index';

/**
 * Context instance that holds the store, orchestrator, and services.
 */
interface ISearchContext {
  store: StoreApi<ISearchStore>;
  orchestrator: SearchOrchestrator;
  managerService: SearchManagerService | undefined;
  isInitialized: boolean;
}

/**
 * Global store registry — maps searchContextId to context instances.
 * Web parts sharing the same searchContextId share the same store.
 */
const contextMap: Map<string, ISearchContext> = new Map();

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
 * Initialize the search context with SPFI for full functionality.
 * Call this once from a web part's onInit() to enable:
 * - Search history logging
 * - Click tracking
 * - Saved searches, collections, etc.
 *
 * @param searchContextId - The shared context ID
 * @param sp - The SPFI instance from PnPjs
 */
export async function initializeSearchContext(
  searchContextId: string,
  sp: SPFI
): Promise<void> {
  const context = getOrCreateContext(searchContextId);

  // Skip if already initialized
  if (context.isInitialized) {
    return;
  }

  // Create and initialize the SearchManagerService
  const managerService = new SearchManagerService(sp);
  await managerService.initialize();
  context.managerService = managerService;

  // Wire the history service to the orchestrator
  context.orchestrator.setHistoryService(managerService);

  // Start the orchestrator (listens to store changes)
  context.orchestrator.start();

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
