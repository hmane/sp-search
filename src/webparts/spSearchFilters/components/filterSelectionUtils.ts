import type { IActiveFilter, IRefinerValue } from '@interfaces/index';
import { normalizeFilterValue } from '@store/utils/filterValueMatching';

export function isRefinerValueSelected(
  filterName: string,
  refinerValue: IRefinerValue,
  activeFilters: IActiveFilter[]
): boolean {
  const refinerRaw = normalizeFilterValue(refinerValue.value);
  const refinerName = normalizeFilterValue(refinerValue.name);

  for (let i = 0; i < activeFilters.length; i++) {
    const active = activeFilters[i];
    if (active.filterName !== filterName) {
      continue;
    }

    const activeRaw = normalizeFilterValue(active.value);
    const activeDisplay = normalizeFilterValue(active.displayValue);

    if (
      active.value === refinerValue.value ||
      activeRaw === refinerRaw ||
      activeRaw === refinerName ||
      (activeDisplay && (activeDisplay === refinerRaw || activeDisplay === refinerName))
    ) {
      return true;
    }
  }

  return false;
}

export function getSelectedRefinerTokens(
  filterName: string,
  values: IRefinerValue[],
  activeFilters: IActiveFilter[]
): string[] {
  const selected: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (isRefinerValueSelected(filterName, values[i], activeFilters)) {
      selected.push(values[i].value);
    }
  }

  for (let i = 0; i < activeFilters.length; i++) {
    const active = activeFilters[i];
    if (active.filterName !== filterName) {
      continue;
    }

    let representedByBucket = false;
    for (let j = 0; j < values.length; j++) {
      if (isRefinerValueSelected(filterName, values[j], [active])) {
        representedByBucket = true;
        break;
      }
    }

    if (!representedByBucket && selected.indexOf(active.value) < 0) {
      selected.push(active.value);
    }
  }

  return selected;
}
