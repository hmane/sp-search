import * as React from 'react';
import type { StoreApi } from 'zustand/vanilla';
import { isInAudience } from '@services/index';
import type { ISearchStore } from '@interfaces/index';

/**
 * Stream D / #10 — per-web-part audience targeting.
 *
 * Wraps a web part's React tree and renders nothing when the current user
 * is not in any of the configured Azure AD groups. Empty `audienceGroups`
 * means "visible to everyone" (today's behaviour).
 *
 * `currentUserGroups` lives in `uiSlice` and is populated fire-and-forget
 * during context init by `storeRegistry.ts` calling `resolveUserGroupIds()`.
 * Until that Graph call resolves, audience-targeted web parts stay hidden
 * — consistent with the fail-closed semantics already in place for
 * verticals + refiners + promoted results.
 *
 * Reactivity: the gate subscribes to the store so it re-evaluates when
 * `currentUserGroups` flips from `[]` to the resolved set — the targeted
 * web parts become visible without a page reload.
 */
export interface IAudienceGateProps {
  audienceGroups: string[];
  /** The shared Zustand store. May be `undefined` during early SPFx mount. */
  store: StoreApi<ISearchStore> | undefined;
  /** Marked optional so callers can pass children as the third arg to `React.createElement`. */
  children?: React.ReactNode;
}

function useCurrentUserGroups(store: StoreApi<ISearchStore> | undefined): string[] {
  const [groups, setGroups] = React.useState<string[]>(function (): string[] {
    return store ? (store.getState().currentUserGroups || []) : [];
  });

  React.useEffect(function (): (() => void) | undefined {
    if (!store) {
      return undefined;
    }
    setGroups(store.getState().currentUserGroups || []);
    const unsubscribe = store.subscribe(function (next: ISearchStore): void {
      const nextGroups = next.currentUserGroups || [];
      setGroups(function (prev: string[]): string[] {
        return prev === nextGroups ? prev : nextGroups;
      });
    });
    return unsubscribe;
  }, [store]);

  return groups;
}

export const AudienceGate: React.FC<IAudienceGateProps> = ({ audienceGroups, store, children }) => {
  const userGroups = useCurrentUserGroups(store);

  if (!audienceGroups || audienceGroups.length === 0) {
    return <>{children}</>;
  }
  if (isInAudience(audienceGroups, userGroups)) {
    return <>{children}</>;
  }
  return null;
};

/**
 * Parse an admin-supplied audience-groups string into a clean array of
 * Azure AD group object IDs. Splits on commas, newlines, semicolons, and
 * whitespace; trims; drops empties. Accepts any of those separators so the
 * admin's choice of paste format doesn't matter.
 */
export function parseAudienceGroups(raw: string | undefined): string[] {
  if (!raw) {
    return [];
  }
  return raw
    .split(/[,;\n\s]+/)
    .map((s: string) => s.trim())
    .filter(Boolean);
}
