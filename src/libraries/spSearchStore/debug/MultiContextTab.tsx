/**
 * T3.D8 — Multi-Context audit panel for the DebugPanel.
 *
 * Enumerates `window.__sp_search_context_map__` and renders one row per
 * search context with: id, urlPrefix (computed + admin override),
 * isInitialized flag, registered provider/action/layout/filter-type
 * counts, the current store snapshot (queryText, vertical,
 * activeFilters count, items.length), urlSyncUnsubscribe status, and
 * registered context IDs from `contextIdRegistry` (which web parts are
 * on which context).
 *
 * "Force dispose" button per row calls `disposeStore(contextId)` after
 * a confirmation. Bypasses the refcount — admin-only debugging tool.
 *
 * Registers via `registerDebugTab('multi-context', ...)` at module
 * load. The DebugPanel consumes the registry and renders the tab
 * alongside its built-in ones (no edit to DebugPanel.tsx needed).
 */

import * as React from 'react';
import type { StoreApi } from 'zustand/vanilla';
import type { ISearchStore } from '@interfaces/index';
import { disposeStore, getContextRefCount } from '@store/store';
import { getRegisteredContextIds } from '../utils/contextIdRegistry';
import { registerDebugTab, type IDebugTabContext } from './debugTabRegistry';

// ─── Window-backed context map shape (read-only, opportunistic) ─────────────

const CONTEXT_MAP_KEY = '__sp_search_context_map__';

interface IContextSummary {
  id: string;
  urlPrefix: string;
  urlPrefixOverride?: string;
  enableUrlSync: boolean;
  isInitialized: boolean;
  urlSyncAttached: boolean;
  store?: StoreApi<ISearchStore>;
  refCount: number;
}

function readContextMap(): IContextSummary[] {
  if (typeof window === 'undefined') { return []; }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const map = (window as any)[CONTEXT_MAP_KEY] as Map<string, {
    store?: StoreApi<ISearchStore>;
    urlSyncUnsubscribe?: () => void;
    isInitialized: boolean;
    urlPrefix: string;
    urlPrefixOverride?: string;
    enableUrlSync: boolean;
  }> | undefined;
  if (!map) { return []; }
  const out: IContextSummary[] = [];
  map.forEach((ctx, id): void => {
    out.push({
      id,
      urlPrefix: ctx.urlPrefix,
      urlPrefixOverride: ctx.urlPrefixOverride,
      enableUrlSync: ctx.enableUrlSync !== false,
      isInitialized: !!ctx.isInitialized,
      urlSyncAttached: !!ctx.urlSyncUnsubscribe,
      store: ctx.store,
      refCount: getContextRefCount(id),
    });
  });
  return out;
}

// ─── Tab renderer ───────────────────────────────────────────────────────────

const MultiContextTabBody: React.FC<{ ctx: IDebugTabContext }> = () => {
  // Periodic refresh — the underlying map updates as web parts mount /
  // unmount; we poll every 2s so the admin sees changes without forcing
  // a Debug Panel re-open.
  const [tick, setTick] = React.useState<number>(0);
  React.useEffect((): (() => void) => {
    const id = window.setInterval((): void => { setTick((n) => n + 1); }, 2000);
    return (): void => { window.clearInterval(id); };
  }, []);

  const contexts = readContextMap();
  const webPartRegistry = getRegisteredContextIds();
  // Invert the web-part → context map into a context → [web parts] map for
  // the registration-source column.
  const webPartsByContext = new Map<string, string[]>();
  webPartRegistry.forEach((contextId, webPartId): void => {
    const arr = webPartsByContext.get(contextId) || [];
    arr.push(webPartId);
    webPartsByContext.set(contextId, arr);
  });

  const handleForceDispose = (contextId: string): void => {
    // eslint-disable-next-line no-alert
    if (window.confirm('Force-dispose context "' + contextId + '"?\n\nThis bypasses the refcount and tears down the store, URL sync, and orchestrator immediately. Any mounted web parts will render their empty/error state until next mount.')) {
      disposeStore(contextId);
      setTick((n) => n + 1);
    }
  };

  if (contexts.length === 0) {
    return (
      <div style={{ padding: 16, color: '#605e5c', fontSize: 13 }}>
        No search contexts currently registered. Mount a search web part to populate this list.
      </div>
    );
  }

  return (
    <div style={{ padding: 12, fontSize: 12 }} data-tick={tick}>
      <div style={{ marginBottom: 8, color: '#605e5c' }}>
        {contexts.length} context{contexts.length === 1 ? '' : 's'} registered. Auto-refreshes every 2s.
      </div>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: 'Consolas, Monaco, monospace' }}>
        <thead>
          <tr style={{ textAlign: 'left', borderBottom: '1px solid #edebe9' }}>
            <th style={{ padding: '6px 8px' }}>Context ID</th>
            <th style={{ padding: '6px 8px' }}>URL Prefix</th>
            <th style={{ padding: '6px 8px', textAlign: 'right' }}>Refs</th>
            <th style={{ padding: '6px 8px' }}>Init</th>
            <th style={{ padding: '6px 8px' }}>URL Sync</th>
            <th style={{ padding: '6px 8px' }}>Web Parts</th>
            <th style={{ padding: '6px 8px' }}>Store snapshot</th>
            <th style={{ padding: '6px 8px' }} />
          </tr>
        </thead>
        <tbody>
          {contexts.map((c): React.ReactElement => {
            const s = c.store ? c.store.getState() : undefined;
            const snapshot = s
              ? 'q=' + JSON.stringify(s.queryText || '') +
                ', vert=' + (s.currentVerticalKey || '') +
                ', filters=' + (s.activeFilters ? s.activeFilters.length : 0) +
                ', items=' + (s.items ? s.items.length : 0)
              : '—';
            const prefix = c.urlPrefixOverride !== undefined
              ? c.urlPrefixOverride + ' (override)'
              : c.urlPrefix || '(none)';
            const webParts = webPartsByContext.get(c.id) || [];
            return (
              <tr key={c.id} style={{ borderBottom: '1px solid #faf9f8' }}>
                <td style={{ padding: '6px 8px', fontWeight: 600 }}>{c.id}</td>
                <td style={{ padding: '6px 8px' }}>{prefix}</td>
                <td style={{ padding: '6px 8px', textAlign: 'right' }}>{c.refCount}</td>
                <td style={{ padding: '6px 8px', color: c.isInitialized ? '#107c10' : '#a4262c' }}>
                  {c.isInitialized ? '✓' : '✗'}
                </td>
                <td style={{ padding: '6px 8px', color: c.urlSyncAttached ? '#107c10' : '#a4262c' }}>
                  {c.enableUrlSync ? (c.urlSyncAttached ? '✓ attached' : '✗ detached') : 'opted out'}
                </td>
                <td style={{ padding: '6px 8px' }}>
                  {webParts.length === 0 ? '—' : webParts.length + ': ' + webParts.slice(0, 3).join(', ') + (webParts.length > 3 ? '...' : '')}
                </td>
                <td style={{ padding: '6px 8px', maxWidth: 280, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={snapshot}>
                  {snapshot}
                </td>
                <td style={{ padding: '6px 8px' }}>
                  <button
                    type="button"
                    onClick={(): void => handleForceDispose(c.id)}
                    style={{
                      padding: '4px 8px',
                      fontSize: 11,
                      backgroundColor: '#a4262c',
                      color: '#ffffff',
                      border: 'none',
                      borderRadius: 2,
                      cursor: 'pointer',
                    }}
                  >
                    Force dispose
                  </button>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
};

// Module-load-time registration. The DebugPanel consumes the registry on
// render so the tab appears the first time the panel opens after this
// module is imported.
registerDebugTab(
  'multi-context',
  'Multi-Context',
  (ctx) => <MultiContextTabBody ctx={ctx} />,
  { sortOrder: 100 }
);

export default MultiContextTabBody;
