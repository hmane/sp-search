/**
 * T3.D7 — canonical dataProviderId validator. Verticals reference data
 * providers by id; the orchestrator's silent fallback (`SearchOrchestrator.ts`)
 * masks misconfiguration — admins typing `'graph'` when they meant
 * `'graph-search'` get back SharePoint Search results in the People
 * vertical with no diagnostic signal. This validator gives the property
 * pane (and the in-component edit-mode MessageBar) a clean signal.
 *
 * Cold-registry semantics: when `registeredIds` is empty (the orchestrator
 * hasn't initialised yet), the validator passes silently. Same shape
 * `validateManagedProperty` uses for a cold schema cache.
 */

import type { IConfigValidationIssue } from './sharedValidators';

/** Cheap edit-distance for Did-You-Mean suggestions. */
function levenshtein(a: string, b: string): number {
  if (!a.length) { return b.length; }
  if (!b.length) { return a.length; }
  const m = a.length;
  const n = b.length;
  const prev: number[] = new Array(n + 1);
  const curr: number[] = new Array(n + 1);
  for (let j = 0; j <= n; j++) { prev[j] = j; }
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a.charAt(i - 1).toLowerCase() === b.charAt(j - 1).toLowerCase() ? 0 : 1;
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    for (let j = 0; j <= n; j++) { prev[j] = curr[j]; }
  }
  return prev[n];
}

function findClosest(target: string, candidates: string[]): string | undefined {
  if (candidates.length === 0) { return undefined; }
  const lower = target.toLowerCase();

  // Prefix match wins — provider ids are kebab-case slugs and typing
  // "graph" when "graph-search" was meant is the common shortening.
  for (let i = 0; i < candidates.length; i++) {
    const cand = candidates[i].toLowerCase();
    if (cand.startsWith(lower) || lower.startsWith(cand)) {
      return candidates[i];
    }
  }

  // Edit-distance fallback for typos like "grpah-search" → "graph-search".
  const threshold = Math.max(3, Math.floor(target.length / 2));
  let bestId: string | undefined;
  let bestDistance = Infinity;
  for (let i = 0; i < candidates.length; i++) {
    const d = levenshtein(target, candidates[i]);
    if (d < bestDistance) {
      bestDistance = d;
      bestId = candidates[i];
    }
  }
  return bestDistance <= threshold ? bestId : undefined;
}

/**
 * Validate one `dataProviderId` against the registry. Returns `''` for
 * valid input; an error message string for invalid. Compatible with the
 * SPFx `onGetErrorMessage` signature.
 */
export function validateDataProviderId(
  value: string | undefined,
  registeredIds: string[]
): string {
  const trimmed = (value || '').trim();
  if (!trimmed) { return ''; }
  if (!registeredIds || registeredIds.length === 0) { return ''; }

  const normalized = trimmed.toLowerCase();
  for (let i = 0; i < registeredIds.length; i++) {
    if (registeredIds[i].toLowerCase() === normalized) {
      return '';
    }
  }

  const suggestion = findClosest(trimmed, registeredIds);
  if (suggestion) {
    return 'Did you mean "' + suggestion + '"? "' + trimmed + '" is not a registered data provider id.';
  }
  return '"' + trimmed + '" is not a registered data provider id. Registered: ' + registeredIds.join(', ') + '.';
}

// ─── Collection-level validator ─────────────────────────────────────────────

export interface IVerticalDataProviderIdRow {
  /** Vertical key (audit signal: error message names the vertical) */
  key: string;
  /** Vertical display label */
  label?: string;
  /** Free-text data provider id from the property pane */
  dataProviderId?: string;
}

/**
 * Validate every vertical's `dataProviderId` against the registry. Returns
 * one `IConfigValidationIssue` per failing row, with `rowIndex` matching the
 * verticalsCollection index so the UI can focus the offending row.
 */
export function validateVerticalDataProviderIds(
  verticals: IVerticalDataProviderIdRow[],
  registeredIds: string[]
): IConfigValidationIssue[] {
  const issues: IConfigValidationIssue[] = [];

  if (!registeredIds || registeredIds.length === 0) {
    return issues;
  }

  for (let i = 0; i < verticals.length; i++) {
    const row = verticals[i];
    const value = (row.dataProviderId || '').trim();
    if (!value) { continue; }

    const message = validateDataProviderId(value, registeredIds);
    if (!message) { continue; }

    const labelDisplay = row.label ? row.label + ' (' + row.key + ')' : row.key;
    issues.push({
      id: 'data-provider-id-' + i,
      severity: 'error',
      rowIndex: i,
      fieldKey: 'dataProviderId',
      message: 'Vertical ' + labelDisplay + ': ' + message,
    });
  }

  return issues;
}
