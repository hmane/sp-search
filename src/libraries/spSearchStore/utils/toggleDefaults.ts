import type { IActiveFilter, IFilterConfig } from '@interfaces/index';

/**
 * Toggle default-value helpers. A `toggle` filter with a `defaultValue` is
 * applied automatically (seeded) but is *implicit* — it is re-seeded on every
 * load/restore, so it should never be persisted to the URL or search history.
 * `seedToggleDefaults` adds those implicit filters; `stripDefaultToggleFilters`
 * removes them. Both are pure (depend only on interfaces) so they can be shared
 * across the store registry, URL middleware, orchestrator, and web parts
 * without creating an import cycle.
 */

/**
 * For each `toggle` filter config that has `defaultValue` set AND whose managed
 * property is not already represented in the current activeFilters, append a
 * synthetic active-filter entry.
 *
 * URL-restore writes to activeFilters BEFORE this runs, so URL state always
 * wins over admin defaults — preserves shareable links.
 *
 * The synthetic entry is structurally identical to one a user would produce by
 * clicking the toggle: `'1'` for true, `'0'` for false, `displayValue` honours
 * trueLabel/falseLabel + invertBoolean, and the operator is taken from the
 * config (matches ToggleFilter.tsx).
 */
export function seedToggleDefaults(
  current: IActiveFilter[],
  configs: IFilterConfig[]
): IActiveFilter[] {
  const activeNames = new Set<string>();
  for (let i = 0; i < current.length; i++) {
    activeNames.add(current[i].filterName);
  }
  const additions: IActiveFilter[] = [];
  for (let i = 0; i < configs.length; i++) {
    const c = configs[i];
    if (c.filterType !== 'toggle') {
      continue;
    }
    if (c.defaultValue === undefined) {
      continue;
    }
    if (activeNames.has(c.managedProperty)) {
      continue;
    }
    const trueLabel = c.trueLabel || 'Yes';
    const falseLabel = c.falseLabel || 'No';
    const invert = c.invertBoolean === true;
    const value = c.defaultValue ? '1' : '0';
    const rawIsTrue = value === '1';
    const displayValue = rawIsTrue
      ? (invert ? falseLabel : trueLabel)
      : (invert ? trueLabel : falseLabel);
    additions.push({
      filterName: c.managedProperty,
      value,
      displayValue,
      operator: c.operator,
    });
  }
  if (additions.length === 0) {
    return current;
  }
  return current.concat(additions);
}

/**
 * Inverse of `seedToggleDefaults`: remove any active filter that sits at its
 * configured toggle default value. Such filters are implicit (auto-seeded on
 * load), so they are excluded from anything that persists or displays user
 * intent — the URL and search history. A toggle overridden away from its
 * default (e.g. the user picks "No") does NOT match and is kept.
 */
export function stripDefaultToggleFilters(
  current: IActiveFilter[],
  configs: IFilterConfig[]
): IActiveFilter[] {
  if (!current || current.length === 0 || !configs || configs.length === 0) {
    return current || [];
  }
  const defaults = new Map<string, string>();
  for (let i = 0; i < configs.length; i++) {
    const c = configs[i];
    if (c.filterType === 'toggle' && c.defaultValue !== undefined) {
      defaults.set(c.managedProperty, c.defaultValue ? '1' : '0');
    }
  }
  if (defaults.size === 0) {
    return current;
  }
  return current.filter((f) => defaults.get(f.filterName) !== f.value);
}
