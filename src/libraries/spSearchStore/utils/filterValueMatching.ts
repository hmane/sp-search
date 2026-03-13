import type { IActiveFilter } from '@interfaces/index';

function stripWrappingQuotes(value: string): string {
  if (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') {
    return value.substring(1, value.length - 1);
  }
  return value;
}

function decodeHexRefinementToken(value: string): string {
  const stripped = stripWrappingQuotes(value);
  if (stripped.indexOf('\u01C2\u01C2') !== 0) {
    return stripped;
  }
  const hex = stripped.substring(2);
  if (!/^[0-9a-fA-F]+$/.test(hex) || hex.length % 2 !== 0) {
    return stripped;
  }
  try {
    let encoded = '';
    for (let i = 0; i < hex.length; i += 2) {
      encoded += '%' + hex.substring(i, i + 2);
    }
    return decodeURIComponent(encoded);
  } catch {
    return stripped;
  }
}

export function normalizeFilterValue(value: string | undefined): string {
  return decodeHexRefinementToken(value || '').trim().toLowerCase();
}

export function areFiltersEquivalent(a: IActiveFilter, b: IActiveFilter): boolean {
  if (a.filterName !== b.filterName) {
    return false;
  }

  const aRaw = normalizeFilterValue(a.value);
  const aDisplay = normalizeFilterValue(a.displayValue);
  const bRaw = normalizeFilterValue(b.value);
  const bDisplay = normalizeFilterValue(b.displayValue);

  if (a.value === b.value) {
    return true;
  }

  return (
    aRaw === bRaw ||
    (aDisplay !== '' && aDisplay === bRaw) ||
    (bDisplay !== '' && aRaw === bDisplay) ||
    (aDisplay !== '' && bDisplay !== '' && aDisplay === bDisplay)
  );
}
