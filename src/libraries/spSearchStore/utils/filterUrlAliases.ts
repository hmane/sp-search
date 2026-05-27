import type { IFilterConfig } from '@interfaces/index';

function normalizeKey(value: string | undefined): string {
  return (value || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
}

export function sanitizeUrlAlias(value: string | undefined): string | undefined {
  const normalized = normalizeKey(value);
  return normalized || undefined;
}

export function getDefaultFilterUrlAlias(
  managedProperty: string,
  filterType?: IFilterConfig['filterType']
): string {
  const normalized = normalizeKey(managedProperty);

  if (filterType === 'people' || normalized === 'authorowsuser' || normalized === 'author') {
    return 'au';
  }
  if (filterType === 'daterange' || normalized === 'lastmodifiedtime' || normalized === 'modified' || normalized === 'modifieddate') {
    return 'md';
  }
  if (normalized === 'filetype') {
    return 'ft';
  }
  if (normalized === 'contenttype') {
    return 'ct';
  }
  if (normalized === 'title') {
    return 'ti';
  }
  if (normalized === 'path') {
    return 'pa';
  }
  if (normalized === 'size') {
    return 'sz';
  }
  if (normalized === 'created') {
    return 'cr';
  }
  if (normalized === 'contentclass') {
    return 'cc';
  }

  return normalized || 'f';
}

export function getFilterUrlAlias(config: Pick<IFilterConfig, 'managedProperty' | 'filterType' | 'urlAlias'>): string {
  return sanitizeUrlAlias(config.urlAlias) || getDefaultFilterUrlAlias(config.managedProperty, config.filterType);
}

/**
 * T3.D3 — Compute disambiguated URL aliases for a full filter-config list.
 *
 * `getFilterUrlAlias` works per-config but can produce collisions when two
 * managed properties resolve to the same default alias (the canonical case
 * is `Author` + `AuthorOWSUSER`, both → `au`). That breaks deep-link
 * round-trip: `?au=value` only restores whichever filter was registered
 * for `au` first.
 *
 * This function walks the configs in order, computes each one's natural
 * alias, and appends a numeric suffix (`au` → `au2` → `au3` …) when a
 * prior filter already claimed the base. The result is a deterministic
 * `Map<filterId, alias>` that callers (urlSyncMiddleware) use for both
 * serialization and deserialization, guaranteeing round-trip consistency.
 *
 * First-come-first-served: the order of the input array determines who
 * keeps the base alias. Admin-supplied `urlAlias` values participate in
 * the collision pool just like defaults — explicit doesn't bypass
 * disambiguation.
 */
export function assignFilterUrlAliases(
  configs: Array<Pick<IFilterConfig, 'id' | 'managedProperty' | 'filterType' | 'urlAlias'>>
): Map<string, string> {
  const result = new Map<string, string>();
  const taken = new Set<string>();

  for (let i: number = 0; i < configs.length; i++) {
    const config = configs[i];
    const base = sanitizeUrlAlias(config.urlAlias) || getDefaultFilterUrlAlias(config.managedProperty, config.filterType);

    let alias = base;
    let suffix = 2;
    while (taken.has(alias)) {
      alias = base + String(suffix);
      suffix++;
    }
    taken.add(alias);
    result.set(config.id, alias);
  }

  return result;
}

