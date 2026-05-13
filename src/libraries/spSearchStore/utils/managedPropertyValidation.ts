import type { IManagedProperty } from '../interfaces/ISearchDataProvider';

/**
 * T4.D3 — synchronous validation for admin-supplied managed-property
 * names against the cached search schema.
 *
 * The audit's acceptance signal calls out two cases:
 *   - "Typo `LastModifedTime` raises did-you-mean before save"
 *   - "Collapse spec rejects non-sortable"
 *
 * Validation runs synchronously off the in-memory / sessionStorage
 * schema cache populated by `SchemaService.fetchManagedProperties`.
 * When the cache is cold the validator passes silently — we never
 * block on a network fetch from inside the property pane.
 */

export interface IValidateOptions {
  /** Require the property to have `sortable: true`. Used by collapseSpecification. */
  requireSortable?: boolean;
  /** Require the property to have `refinable: true`. Used by refiner managed-property fields. */
  requireRefinable?: boolean;
  /** Require the property to have `retrievable: true`. Used by selectedProperties / column fields. */
  requireRetrievable?: boolean;
}

export type ValidationResult =
  | { valid: true }
  | { valid: false; message: string };

/** Cheap edit-distance — sufficient for typo detection in short identifiers. */
function levenshtein(a: string, b: string): number {
  if (!a.length) { return b.length; }
  if (!b.length) { return a.length; }
  const m = a.length;
  const n = b.length;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const prev: number[] = new Array(n + 1);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const curr: number[] = new Array(n + 1);
  for (let j = 0; j <= n; j++) { prev[j] = j; }
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a.charAt(i - 1).toLowerCase() === b.charAt(j - 1).toLowerCase() ? 0 : 1;
      curr[j] = Math.min(
        curr[j - 1] + 1,        // insertion
        prev[j] + 1,            // deletion
        prev[j - 1] + cost,     // substitution
      );
    }
    for (let j = 0; j <= n; j++) { prev[j] = curr[j]; }
  }
  return prev[n];
}

function findExact(value: string, schema: IManagedProperty[]): IManagedProperty | undefined {
  const target = value.toLowerCase();
  for (let i: number = 0; i < schema.length; i++) {
    if (schema[i].name.toLowerCase() === target) {
      return schema[i];
    }
  }
  return undefined;
}

/**
 * Return the closest schema property name to `input` (within an edit
 * distance threshold), or `undefined` if no near match exists or the
 * input already matches a schema entry.
 */
export function suggestCloseManagedProperty(
  input: string,
  schema: IManagedProperty[] | undefined
): string | undefined {
  if (!input || !schema || schema.length === 0) {
    return undefined;
  }
  if (findExact(input, schema)) {
    return undefined;
  }

  // Edit-distance threshold scales with input length — short identifiers
  // tolerate fewer typos than long ones.
  const threshold = Math.max(2, Math.floor(input.length / 4));
  let bestName: string | undefined;
  let bestDistance = Infinity;
  for (let i: number = 0; i < schema.length; i++) {
    const distance = levenshtein(input, schema[i].name);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestName = schema[i].name;
    }
  }
  if (bestDistance <= threshold) {
    return bestName;
  }
  return undefined;
}

/**
 * Validate a managed-property name against the cached schema. Empty
 * input passes (caller decides if the field is required). Empty schema
 * passes (cache cold — don't block the admin on a fetch we can't run
 * synchronously).
 */
export function validateManagedProperty(
  value: string,
  schema: IManagedProperty[] | undefined,
  options: IValidateOptions = {}
): ValidationResult {
  const trimmed = (value || '').trim();
  if (!trimmed) {
    return { valid: true };
  }
  if (!schema || schema.length === 0) {
    return { valid: true };
  }

  const exact = findExact(trimmed, schema);
  if (!exact) {
    const suggestion = suggestCloseManagedProperty(trimmed, schema);
    if (suggestion) {
      return { valid: false, message: 'Did you mean "' + suggestion + '"? "' + trimmed + '" is not a known managed property on this tenant.' };
    }
    return { valid: false, message: '"' + trimmed + '" is not a known managed property on this tenant. Check the spelling or open Search admin → Schema → Managed Properties.' };
  }

  if (options.requireSortable && !exact.sortable) {
    return { valid: false, message: '"' + exact.name + '" is not sortable. Mark it Sortable on the Search admin schema page before using it as a collapse specification.' };
  }
  if (options.requireRefinable && !exact.refinable) {
    return { valid: false, message: '"' + exact.name + '" is not refinable. Mark it Refinable on the Search admin schema page before using it as a refiner.' };
  }
  if (options.requireRetrievable && !exact.retrievable) {
    return { valid: false, message: '"' + exact.name + '" is not retrievable. Mark it Retrievable on the Search admin schema page before requesting it as a selected property.' };
  }

  return { valid: true };
}
