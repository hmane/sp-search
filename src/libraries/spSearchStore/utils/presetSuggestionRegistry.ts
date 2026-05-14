/**
 * T4.D12 — cross-web-part scenario preset propagation.
 *
 * When the Results web part applies a scenario preset (`_applyScenarioPreset`)
 * the preset carries `filterSuggestions` that the Results web part itself
 * doesn't own — those filters configure the Filters web part. This registry
 * is the carrier: Results records the latest preset suggestion under the
 * page's `searchContextId`, and Filters subscribes in edit mode to surface
 * a MessageBar offering to apply the suggested filters.
 *
 * Window-backed like `contextIdRegistry` so each web part bundle's separate
 * webpack entry sees the same Map.
 */

import type { IPresetFilterSuggestion } from '../../../webparts/spSearchResults/presets/searchPresets';

export interface IPresetSuggestion {
  /** Preset id (e.g. 'documents', 'people') — matches `IScenarioPreset.id`. */
  id: string;
  /** Preset label for display in the MessageBar (e.g. 'Documents'). */
  label: string;
  /** Filters the preset suggests for the Filters web part. */
  filterSuggestions: IPresetFilterSuggestion[];
  /** Timestamp when the suggestion was recorded — set by `recordPresetSuggestion`. */
  recordedAt: number;
}

const REGISTRY_KEY = '__sp_search_preset_suggestion_registry_v1__';

interface IRegistryEntry {
  /** Map<searchContextId, IPresetSuggestion> */
  suggestions: Map<string, IPresetSuggestion>;
  /** Subscribers fire on every record/clear. */
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
    entry = { suggestions: new Map<string, IPresetSuggestion>(), listeners: new Set<() => void>() };
    win[REGISTRY_KEY] = entry;
  }
  return entry;
}

function notify(): void {
  const { listeners } = getRegistry();
  listeners.forEach((listener) => {
    try { listener(); } catch { /* swallow — one bad subscriber shouldn't break others */ }
  });
}

/**
 * Record a preset suggestion for a search context. Overwrites any previous
 * suggestion for the same context. Fires every subscriber.
 */
export function recordPresetSuggestion(searchContextId: string, suggestion: IPresetSuggestion): void {
  const { suggestions } = getRegistry();
  suggestions.set(searchContextId, {
    id: suggestion.id,
    label: suggestion.label,
    filterSuggestions: suggestion.filterSuggestions,
    recordedAt: Date.now(),
  });
  notify();
}

/**
 * Read (peek) the current suggestion for a context. Does NOT remove it —
 * the caller decides via `clearPresetSuggestion` when to drop it
 * (typically on Apply / Dismiss).
 */
export function consumePresetSuggestion(searchContextId: string): IPresetSuggestion | undefined {
  return getRegistry().suggestions.get(searchContextId);
}

/**
 * Remove the suggestion for a context. No-ops when nothing is stored.
 * Fires subscribers so re-render can hide the MessageBar.
 */
export function clearPresetSuggestion(searchContextId: string): void {
  const { suggestions } = getRegistry();
  if (suggestions.has(searchContextId)) {
    suggestions.delete(searchContextId);
    notify();
  }
}

/**
 * Subscribe to record/clear events. Returns an unsubscribe function.
 */
export function subscribePresetSuggestionChanges(listener: () => void): () => void {
  const { listeners } = getRegistry();
  listeners.add(listener);
  return (): void => { listeners.delete(listener); };
}

/** Test-only — clears the entire registry. */
export function _resetPresetSuggestionRegistryForTesting(): void {
  const entry = getRegistry();
  entry.suggestions.clear();
  entry.listeners.clear();
}
