/**
 * T3.D2 — window-backed cross-web-part registry of which
 * `searchContextId` each SPFx web part instance is currently using.
 *
 * Each web part bundle (Box / Results / Filters / Verticals / Manager /
 * AdminManager) is a separate webpack entry with its own module-level
 * state, so a per-bundle Map would not see across web parts. The registry
 * is parked on `window` under a versioned key so every bundle sees the
 * same Map. This mirrors the pattern `storeRegistry.ts` already uses for
 * `window.__sp_search_context_map__`.
 *
 * Usage:
 *   - Each web part calls `setWebPartContextId(webPartId, contextId)` on
 *     mount and on every `onPropertyPaneFieldChanged('searchContextId')`.
 *   - Each web part calls `unregisterWebPartContextId(webPartId)` on
 *     `onDispose`.
 *   - The mismatch banner reads `getRegisteredContextIds()` to compute
 *     the set of IDs in play.
 *
 * The registry never invokes React; consumers subscribe via the
 * `subscribeContextIdChanges` callback so they can re-render when peer
 * web parts come, go, or change their context.
 */

const REGISTRY_KEY = '__sp_search_context_id_registry_v1__';

interface IRegistryEntry {
  /** Map<webPartId, contextId>. */
  ids: Map<string, string>;
  /** Subscribers — fire on every change. */
  listeners: Set<() => void>;
}

interface IWindowWithRegistry {
  [REGISTRY_KEY]?: IRegistryEntry;
}

function getRegistry(): IRegistryEntry {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const win = (typeof window !== 'undefined' ? window : ({} as any)) as IWindowWithRegistry;
  let entry = win[REGISTRY_KEY];
  if (!entry) {
    entry = { ids: new Map<string, string>(), listeners: new Set<() => void>() };
    win[REGISTRY_KEY] = entry;
  }
  return entry;
}

function notify(): void {
  const { listeners } = getRegistry();
  listeners.forEach(function (listener: () => void): void {
    try {
      listener();
    } catch {
      // Listeners must not throw — swallow per design.
    }
  });
}

export function setWebPartContextId(webPartId: string, contextId: string): void {
  if (!webPartId) {
    return;
  }
  const { ids } = getRegistry();
  const previous = ids.get(webPartId);
  if (previous === contextId) {
    return;
  }
  ids.set(webPartId, contextId);
  notify();
}

export function unregisterWebPartContextId(webPartId: string): void {
  if (!webPartId) {
    return;
  }
  const { ids } = getRegistry();
  if (ids.has(webPartId)) {
    ids.delete(webPartId);
    notify();
  }
}

export function getRegisteredContextIds(): Map<string, string> {
  // Return a defensive copy so callers can't mutate the source.
  return new Map(getRegistry().ids);
}

export function subscribeContextIdChanges(listener: () => void): () => void {
  const { listeners } = getRegistry();
  listeners.add(listener);
  return function unsubscribe(): void {
    listeners.delete(listener);
  };
}

/**
 * T3.D2 — compute whether the given web part is in a mismatched state
 * against its peers. Returns the set of OTHER context IDs currently
 * registered (excluding `thisWebPartId`'s own). Empty set means there
 * are no peers or peers all match this one.
 */
export function getPeerContextIds(thisWebPartId: string, thisContextId: string): string[] {
  const { ids } = getRegistry();
  const peers = new Set<string>();
  ids.forEach(function (contextId: string, webPartId: string): void {
    if (webPartId !== thisWebPartId && contextId !== thisContextId) {
      peers.add(contextId);
    }
  });
  return Array.from(peers);
}

/** Test-only: clear the registry between specs. */
export function _resetContextIdRegistryForTesting(): void {
  const entry = getRegistry();
  entry.ids.clear();
  entry.listeners.clear();
}
