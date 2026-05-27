/**
 * T5.D4 — extensible DebugPanel tab registry.
 *
 * Cross-track contributions (e.g. T3.D8's Multi-Context audit panel)
 * register a tab via `registerDebugTab(id, label, render)` instead of
 * editing `DebugPanel.tsx`. The panel renders built-in tabs from its
 * static TABS array first, then appends registered tabs in
 * `sortOrder` ascending then registration order.
 *
 * Per audit cross-track contract: T3.D8 registers `id='multi-context'`
 * via this API. The registry is window-backed so cross-bundle imports
 * see the same list of registered tabs.
 */

import * as React from 'react';
import type { StoreApi } from 'zustand/vanilla';
import type { ISearchStore } from '@interfaces/index';
import { DebugCollector } from './DebugCollector';

/**
 * Context passed to a registered tab's `render` function. Contributed
 * tabs read from these shared data sources without taking a hard
 * dependency on `DebugPanel`'s internals.
 */
export interface IDebugTabContext {
  /** Store of the owning web part (may be undefined in edge cases). */
  store?: StoreApi<ISearchStore>;
  /** DebugCollector singleton — surfaces buffered events. */
  debugCollector: typeof DebugCollector;
}

export interface IDebugTabRegistration {
  id: string;
  label: string;
  render: (ctx: IDebugTabContext) => React.ReactElement;
  sortOrder: number;
}

const REGISTRY_KEY = '__sp_search_debug_tab_registry_v1__';

interface IRegistryEntry {
  /** Map<tab id, registration> — last-write-wins on re-register. */
  tabs: Map<string, IDebugTabRegistration>;
  /** Monotonic counter used as the implicit secondary sort. */
  nextOrder: number;
}

interface IWindowWithRegistry {
  [REGISTRY_KEY]?: IRegistryEntry;
}

function getRegistry(): IRegistryEntry {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const win = (typeof window !== 'undefined' ? window : ({} as any)) as IWindowWithRegistry;
  let entry = win[REGISTRY_KEY];
  if (!entry) {
    entry = { tabs: new Map<string, IDebugTabRegistration>(), nextOrder: 1 };
    win[REGISTRY_KEY] = entry;
  }
  return entry;
}

export interface IRegisterDebugTabOptions {
  /** Lower numbers render first. Default: monotonic counter (registration order). */
  sortOrder?: number;
}

/**
 * Register (or replace) a DebugPanel tab. Tabs persist for the lifetime
 * of the page; calling with the same `id` overwrites the previous entry.
 * The DebugPanel re-reads the registry on every render so newly-
 * registered tabs appear without a panel re-mount.
 */
export function registerDebugTab(
  id: string,
  label: string,
  render: (ctx: IDebugTabContext) => React.ReactElement,
  options?: IRegisterDebugTabOptions
): void {
  const reg = getRegistry();
  const sortOrder = options && typeof options.sortOrder === 'number' ? options.sortOrder : reg.nextOrder++;
  reg.tabs.set(id, { id, label, render, sortOrder });
}

/**
 * Return the registered tabs sorted by `sortOrder` ascending then
 * registration order. The DebugPanel appends them after its built-in
 * tabs.
 */
export function getRegisteredDebugTabs(): IDebugTabRegistration[] {
  const list = Array.from(getRegistry().tabs.values());
  list.sort((a, b) => a.sortOrder - b.sortOrder);
  return list;
}

/**
 * Remove a registered tab. No-op if the id isn't registered.
 */
export function unregisterDebugTab(id: string): void {
  getRegistry().tabs.delete(id);
}

/**
 * Test-only — clear the entire registry.
 */
export function _resetDebugTabRegistryForTesting(): void {
  const reg = getRegistry();
  reg.tabs.clear();
  reg.nextOrder = 1;
}
