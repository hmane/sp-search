export { createSearchStore } from './createStore';
export {
  getStore,
  disposeStore,
  hasStore,
  getOrchestrator,
  getManagerService,
  initializeSearchContext,
  // T3.D1 — refcounted dispose contract for web part lifecycles.
  incrementContextRef,
  decrementContextRef,
  getContextRefCount
} from './storeRegistry';
export type { IInitializeContextOptions } from './storeRegistry';
export {
  serializeToUrl,
  deserializeFromUrl,
  createUrlSyncSubscription
} from './middleware';
