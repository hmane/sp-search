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

