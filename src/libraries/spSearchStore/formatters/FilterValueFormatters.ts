import 'spfx-toolkit/lib/utilities/context/pnpImports/taxonomy';
import type { IFilterConfig, IFilterValueFormatter } from '@interfaces/index';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';

interface ITaxonomyLabel {
  label: string;
  path: string;
}

interface INumericRange {
  min?: number;
  max?: number;
}

const taxonomyCache: Map<string, Promise<ITaxonomyLabel>> = new Map();
const peopleCache: Map<string, Promise<string>> = new Map();

function stripStringWrapper(value: string): string {
  if (value.indexOf('string("') === 0 && value.lastIndexOf('")') === value.length - 2) {
    return value.substring(8, value.length - 2);
  }
  return value;
}

function encodeUrlValue(value: string): string {
  return encodeURIComponent(value);
}

function decodeUrlValue(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function extractGuid(value: string): string | undefined {
  const match = value.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
  return match ? match[0] : undefined;
}

function extractEmailFromClaim(claim: string): string {
  const parts = claim.split('|');
  const last = parts[parts.length - 1];
  return last || claim;
}

function getDefaultLabel(labels: Array<{ name: string; isDefault: boolean }> | undefined): string {
  if (!labels || labels.length === 0) {
    return '';
  }
  const found = labels.find((label) => label.isDefault);
  return found ? found.name : labels[0].name;
}

async function resolveTaxonomyLabel(termId: string): Promise<ITaxonomyLabel> {
  const cached = taxonomyCache.get(termId);
  if (cached) {
    return cached;
  }

  const promise = (async (): Promise<ITaxonomyLabel> => {
    try {
      const term = await (SPContext.sp as any).termStore.getTermById(termId)();
      const label = getDefaultLabel(term?.labels);
      if (term && term.parent && term.parent.id) {
        const parentLabel = await resolveTaxonomyLabel(term.parent.id as string);
        return {
          label: label || termId,
          path: parentLabel.path + ' > ' + (label || termId),
        };
      }
      return {
        label: label || termId,
        path: label || termId,
      };
    } catch (error) {
      SPContext.logger.warn('Taxonomy label resolution failed', { termId, error });
      return {
        label: '(Unknown term)',
        path: '(Unknown term)',
      };
    }
  })();

  taxonomyCache.set(termId, promise);
  return promise;
}

async function resolvePeopleDisplayName(claim: string): Promise<string> {
  const cached = peopleCache.get(claim);
  if (cached) {
    return cached;
  }

  const promise = (async (): Promise<string> => {
    try {
      const profile = await (SPContext.sp as any).profiles.getPropertiesFor(claim);
      if (profile && profile.DisplayName) {
        return profile.DisplayName as string;
      }
      const properties = profile?.UserProfileProperties as Array<{ Key: string; Value: string }> | undefined;
      if (properties) {
        const preferred = properties.find((prop) => prop.Key === 'PreferredName');
        if (preferred && preferred.Value) {
          return preferred.Value;
        }
      }
    } catch (error) {
      SPContext.logger.warn('People profile resolution failed', { claim, error });
    }
    return extractEmailFromClaim(claim);
  })();

  peopleCache.set(claim, promise);
  return promise;
}

function parseRangePart(value: string): number | undefined {
  const trimmed = value.trim();
  if (trimmed === 'max') {
    return undefined;
  }
  const decimalMatch = trimmed.match(/decimal\(([^)]+)\)/i);
  const candidate = decimalMatch ? decimalMatch[1] : trimmed;
  const parsed = parseFloat(candidate);
  return isNaN(parsed) ? undefined : parsed;
}

export function parseRangeToken(rawValue: string): INumericRange | undefined {
  const value = stripStringWrapper(rawValue);
  if (value.indexOf('range(') !== 0 || value[value.length - 1] !== ')') {
    return undefined;
  }
  const inner = value.substring(6, value.length - 1);
  const parts = inner.split(',');
  if (parts.length < 2) {
    return undefined;
  }
  const min = parseRangePart(parts[0]);
  const max = parseRangePart(parts[1]);
  return { min, max };
}

function formatBytes(value: number): string {
  if (value < 1024) {
    return String(Math.round(value)) + ' B';
  }
  const units = ['KB', 'MB', 'GB', 'TB'];
  let size = value / 1024;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size = size / 1024;
    unitIndex++;
  }
  const rounded = size < 10 ? size.toFixed(1) : Math.round(size).toString();
  return rounded + ' ' + units[unitIndex];
}

export function formatNumericValue(value: number, config: IFilterConfig): string {
  const format = config.rangeFormat || 'number';
  if (format === 'bytes') {
    return formatBytes(value);
  }
  if (format === 'currency') {
    const currency = config.currency || 'USD';
    try {
      return new Intl.NumberFormat(undefined, { style: 'currency', currency }).format(value);
    } catch {
      return currency + ' ' + value.toFixed(0);
    }
  }
  return value.toLocaleString();
}

function formatNumericRange(range: INumericRange, config: IFilterConfig): string {
  if (range.min === undefined && range.max === undefined) {
    return '';
  }
  if (range.min !== undefined && range.max !== undefined) {
    return formatNumericValue(range.min, config) + ' - ' + formatNumericValue(range.max, config);
  }
  if (range.min !== undefined) {
    return '> ' + formatNumericValue(range.min, config);
  }
  return '< ' + formatNumericValue(range.max as number, config);
}

function buildNumericRangeToken(range: INumericRange): string {
  const min = range.min !== undefined ? 'decimal(' + String(range.min) + ')' : 'min';
  const max = range.max !== undefined ? 'decimal(' + String(range.max) + ')' : 'max';
  return 'range(' + min + ', ' + max + ')';
}

export const DefaultFilterFormatter: IFilterValueFormatter = {
  id: 'default',
  formatForDisplay: (rawValue: string): string => stripStringWrapper(rawValue),
  formatForQuery: (displayValue: unknown): string => String(displayValue || ''),
  formatForUrl: (rawValue: string): string => encodeUrlValue(rawValue),
  parseFromUrl: (urlValue: string): string => decodeUrlValue(urlValue),
};

export const TaxonomyFilterFormatter: IFilterValueFormatter = {
  id: 'taxonomy',
  formatForDisplay: async (rawValue: string): Promise<string> => {
    const stripped = stripStringWrapper(rawValue);
    const guid = extractGuid(stripped);
    if (!guid) {
      return stripped;
    }
    const resolved = await resolveTaxonomyLabel(guid);
    return resolved.path || resolved.label || stripped;
  },
  formatForQuery: (displayValue: unknown): string => {
    if (typeof displayValue === 'string') {
      if (displayValue.indexOf('GP0|#') === 0) {
        return displayValue;
      }
      const guid = extractGuid(displayValue);
      if (guid) {
        return 'GP0|#' + guid;
      }
      return displayValue;
    }
    const value = displayValue as { id?: string; termId?: string } | undefined;
    const id = value?.termId || value?.id;
    return id ? 'GP0|#' + id : '';
  },
  formatForUrl: (rawValue: string): string => encodeUrlValue(rawValue),
  parseFromUrl: (urlValue: string): string => decodeUrlValue(urlValue),
};

export const PeopleFilterFormatter: IFilterValueFormatter = {
  id: 'people',
  formatForDisplay: async (rawValue: string): Promise<string> => {
    const stripped = stripStringWrapper(rawValue);
    if (stripped.indexOf('i:0#.f|') !== 0 && stripped.indexOf('|') < 0) {
      return stripped;
    }
    return resolvePeopleDisplayName(stripped);
  },
  formatForQuery: (displayValue: unknown): string => {
    if (typeof displayValue === 'string') {
      if (displayValue.indexOf('|') >= 0) {
        return displayValue;
      }
      return 'i:0#.f|membership|' + displayValue;
    }
    return '';
  },
  formatForUrl: (rawValue: string): string => encodeUrlValue(rawValue),
  parseFromUrl: (urlValue: string): string => decodeUrlValue(urlValue),
};

export const NumericFilterFormatter: IFilterValueFormatter = {
  id: 'slider',
  formatForDisplay: (rawValue: string, config: IFilterConfig): string => {
    const range = parseRangeToken(rawValue);
    if (!range) {
      return stripStringWrapper(rawValue);
    }
    return formatNumericRange(range, config);
  },
  formatForQuery: (displayValue: unknown): string => {
    if (typeof displayValue === 'string') {
      return displayValue;
    }
    if (Array.isArray(displayValue)) {
      return buildNumericRangeToken({ min: displayValue[0], max: displayValue[1] });
    }
    const range = displayValue as INumericRange | undefined;
    if (!range) {
      return '';
    }
    return buildNumericRangeToken(range);
  },
  formatForUrl: (rawValue: string): string => encodeUrlValue(rawValue),
  parseFromUrl: (urlValue: string): string => decodeUrlValue(urlValue),
};

function parseDateRangePart(value: string): string | undefined {
  const match = value.match(/datetime\("([^"]+)"\)/i);
  return match ? match[1] : undefined;
}

function parseDateRangeToken(rawValue: string): { from: Date; to: Date } | undefined {
  const value = stripStringWrapper(rawValue);
  if (value.indexOf('range(') !== 0 || value[value.length - 1] !== ')') {
    return undefined;
  }
  const inner = value.substring(6, value.length - 1);
  const commaIdx = inner.indexOf(',');
  if (commaIdx < 0) {
    return undefined;
  }
  const fromStr = parseDateRangePart(inner.substring(0, commaIdx));
  const toStr = parseDateRangePart(inner.substring(commaIdx + 1));
  if (!fromStr || !toStr) {
    return undefined;
  }
  const from = new Date(fromStr);
  const to = new Date(toStr);
  if (isNaN(from.getTime()) || isNaN(to.getTime())) {
    return undefined;
  }
  return { from, to };
}

function formatShortDate(date: Date): string {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

function buildDateRangeToken(from: Date, to: Date): string {
  return 'range(datetime("' + from.toISOString() + '"), datetime("' + to.toISOString() + '"))';
}

export const DateFilterFormatter: IFilterValueFormatter = {
  id: 'daterange',
  formatForDisplay: (rawValue: string): string => {
    const parsed = parseDateRangeToken(rawValue);
    if (!parsed) {
      return stripStringWrapper(rawValue);
    }
    return formatShortDate(parsed.from) + ' \u2013 ' + formatShortDate(parsed.to);
  },
  formatForQuery: (displayValue: unknown): string => {
    if (typeof displayValue === 'string') {
      return displayValue;
    }
    if (Array.isArray(displayValue) && displayValue.length >= 2) {
      return buildDateRangeToken(new Date(displayValue[0]), new Date(displayValue[1]));
    }
    return '';
  },
  formatForUrl: (rawValue: string): string => encodeUrlValue(rawValue),
  parseFromUrl: (urlValue: string): string => decodeUrlValue(urlValue),
};

export const BooleanFilterFormatter: IFilterValueFormatter = {
  id: 'toggle',
  formatForDisplay: (rawValue: string, config: IFilterConfig): string => {
    const stripped = stripStringWrapper(rawValue);
    if (stripped === '1' || stripped.toLowerCase() === 'true') {
      return config.trueLabel || 'Yes';
    }
    if (stripped === '0' || stripped.toLowerCase() === 'false') {
      return config.falseLabel || 'No';
    }
    return stripped;
  },
  formatForQuery: (displayValue: unknown, config: IFilterConfig): string => {
    if (typeof displayValue === 'boolean') {
      return displayValue ? '1' : '0';
    }
    if (typeof displayValue === 'string') {
      const normalized = displayValue.toLowerCase();
      if (normalized === 'yes' || normalized === (config.trueLabel || '').toLowerCase()) {
        return '1';
      }
      if (normalized === 'no' || normalized === (config.falseLabel || '').toLowerCase()) {
        return '0';
      }
      if (normalized === '1' || normalized === '0') {
        return normalized;
      }
    }
    return '';
  },
  formatForUrl: (rawValue: string): string => encodeUrlValue(rawValue),
  parseFromUrl: (urlValue: string): string => decodeUrlValue(urlValue),
};

const formatterRegistry: Map<string, IFilterValueFormatter> = new Map([
  ['taxonomy', TaxonomyFilterFormatter],
  ['people', PeopleFilterFormatter],
  ['slider', NumericFilterFormatter],
  ['toggle', BooleanFilterFormatter],
  ['daterange', DateFilterFormatter],
  ['checkbox', DefaultFilterFormatter],
  ['tagbox', DefaultFilterFormatter],
  ['default', DefaultFilterFormatter],
]);

export function getFilterValueFormatter(filterType: string | undefined): IFilterValueFormatter {
  if (!filterType) {
    return DefaultFilterFormatter;
  }
  return formatterRegistry.get(filterType) || DefaultFilterFormatter;
}
