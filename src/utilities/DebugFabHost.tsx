/**
 * T5.D1 — cross-bundle singleton FAB host.
 *
 * Each user-facing web part mounts `<DebugFabHost store={...} />` and
 * the first one to render claims the window-backed owner flag. Only
 * the owner renders the FAB + DebugPanel pair; non-owners render
 * nothing. On unmount, the owner releases the flag — the next
 * remaining web part claims it on its next render cycle (the polling
 * effect retries periodically when not the owner).
 *
 * The audit's acceptance signal: "Box-only page with ?debug=1 renders
 * FAB on Box; clicking opens singleton DebugPanel; 6-web-part page
 * renders one FAB."
 */

import * as React from 'react';
import type { StoreApi } from 'zustand/vanilla';
import type { ISearchStore } from '@interfaces/index';
import { DebugCollector } from '@store/debug';

// Lazy-loaded — DebugFab + DebugPanel only land in the bundle when
// the user activates the debug surface (?debug=1).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const DebugFab: any = React.lazy(
  () => import(/* webpackChunkName: 'DebugFab' */ '../webparts/spSearchResults/components/DebugFab') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>
);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const DebugPanel: any = React.lazy(
  () => import(/* webpackChunkName: 'DebugPanel' */ '../webparts/spSearchResults/components/DebugPanel') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>
);

const OWNER_KEY = '__sp_search_debug_fab_owner__';
const CLAIM_POLL_MS = 500;

interface IWindowWithOwner {
  [OWNER_KEY]?: string;
}

function tryClaim(instanceId: string): boolean {
  if (typeof window === 'undefined') { return false; }
  const win = window as unknown as IWindowWithOwner;
  if (!win[OWNER_KEY] || win[OWNER_KEY] === instanceId) {
    win[OWNER_KEY] = instanceId;
    return true;
  }
  return false;
}

function release(instanceId: string): void {
  if (typeof window === 'undefined') { return; }
  const win = window as unknown as IWindowWithOwner;
  if (win[OWNER_KEY] === instanceId) {
    win[OWNER_KEY] = undefined;
  }
}

export interface IDebugFabHostProps {
  /**
   * Store for the DebugPanel's State tab. Every shipped web part has a
   * store via `getStore(searchContextId)` so this is required in
   * practice; the type-safer alternative would be branching the
   * DebugPanel rendering on undefined, which 500+ lines of existing
   * code doesn't currently support.
   */
  store: StoreApi<ISearchStore>;
}

export const DebugFabHost: React.FC<IDebugFabHostProps> = ({ store }) => {
  const instanceIdRef = React.useRef<string>('fab-' + Math.random().toString(36).substring(2));
  const [isOwner, setIsOwner] = React.useState<boolean>(() => tryClaim(instanceIdRef.current));
  const [debugOpen, setDebugOpen] = React.useState<boolean>(false);

  // Poll for ownership when we're not the owner. The current owner may
  // release the flag (e.g. its web part unmounts during a SPA navigation);
  // a low-frequency poll lets the next mounted web part take over.
  React.useEffect((): (() => void) | undefined => {
    if (isOwner) { return undefined; }
    const intervalId = window.setInterval((): void => {
      if (tryClaim(instanceIdRef.current)) {
        setIsOwner(true);
      }
    }, CLAIM_POLL_MS);
    return (): void => { window.clearInterval(intervalId); };
  }, [isOwner]);

  // Release the flag on unmount.
  React.useEffect((): (() => void) => {
    const id = instanceIdRef.current;
    return (): void => { release(id); };
  }, []);

  if (!isOwner) { return null; }

  const isDebugActive = DebugCollector.isActive();
  if (!isDebugActive) { return null; }

  return (
    <React.Suspense fallback={null}>
      {!debugOpen && (
        <DebugFab onClick={(): void => setDebugOpen(true)} />
      )}
      {debugOpen && (
        <DebugPanel store={store} onClose={(): void => setDebugOpen(false)} />
      )}
    </React.Suspense>
  );
};

export default DebugFabHost;
