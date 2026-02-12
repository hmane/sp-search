import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IActiveFilter, IFilterConfig } from '@interfaces/index';
import { getFilterValueFormatter } from '@store/formatters/FilterValueFormatters';
import styles from './SpSearchResults.module.scss';

export interface IActiveFilterPillBarProps {
  activeFilters: IActiveFilter[];
  filterConfig: IFilterConfig[];
  onRemoveFilter: (filterName: string, value?: string) => void;
  onClearAll: () => void;
  onReset: () => void;
}

/**
 * Resolves a managed property name to a human-readable display name
 * using the filter config array.
 */
function getDisplayName(filterName: string, filterConfig: IFilterConfig[]): string {
  for (let i = 0; i < filterConfig.length; i++) {
    if (filterConfig[i].managedProperty === filterName) {
      return filterConfig[i].displayName;
    }
  }
  return filterName;
}

function getFilterConfig(filterName: string, filterConfig: IFilterConfig[]): IFilterConfig {
  for (let i = 0; i < filterConfig.length; i++) {
    if (filterConfig[i].managedProperty === filterName) {
      return filterConfig[i];
    }
  }
  return {
    id: filterName,
    displayName: filterName,
    managedProperty: filterName,
    filterType: 'checkbox',
    operator: 'OR',
    maxValues: 0,
    defaultExpanded: true,
    showCount: true,
    sortBy: 'count',
    sortDirection: 'desc',
    multiValues: true,
  };
}

/**
 * Group active filters by filterName so multi-value filters are
 * combined into one pill with comma-separated values.
 */
function groupFiltersByName(
  filters: IActiveFilter[]
): Map<string, Array<{ value: string; displayValue?: string }>> {
  const groups = new Map<string, Array<{ value: string; displayValue?: string }>>();
  for (let i = 0; i < filters.length; i++) {
    const filter = filters[i];
    const entry = { value: filter.value, displayValue: filter.displayValue };
    const existing = groups.get(filter.filterName);
    if (existing) {
      existing.push(entry);
    } else {
      groups.set(filter.filterName, [entry]);
    }
  }
  return groups;
}

/**
 * Decode a SharePoint hex-encoded refinement token (ǂǂ prefix) to plain text.
 * E.g. "ǂǂ646f6378" → "docx", ǂǂ706466 → "pdf"
 */
function decodeHexToken(token: string): string | undefined {
  // Strip surrounding quotes if present
  let inner = token;
  if (inner.charAt(0) === '"' && inner.charAt(inner.length - 1) === '"') {
    inner = inner.substring(1, inner.length - 1);
  }
  // Check for ǂǂ prefix (U+01C2 = Latin Letter Alveolar Click)
  if (inner.indexOf('\u01C2\u01C2') !== 0) {
    return undefined;
  }
  const hex = inner.substring(2);
  if (hex.length === 0 || hex.length % 2 !== 0) {
    return undefined;
  }
  try {
    const bytes: number[] = [];
    for (let i = 0; i < hex.length; i += 2) {
      bytes.push(parseInt(hex.substring(i, i + 2), 16));
    }
    return new TextDecoder().decode(new Uint8Array(bytes));
  } catch {
    return undefined;
  }
}

/**
 * Format a raw refiner token value for human-readable display.
 * Decodes hex tokens, strips FQL wrappers, GUID prefixes, date ranges, etc.
 */
function formatValueForDisplay(rawValue: string): string {
  // Decode hex-encoded refinement tokens (ǂǂ prefix) — common after page refresh
  const decoded = decodeHexToken(rawValue);
  if (decoded) {
    return decoded;
  }
  // Strip FQL string() wrapper
  if (rawValue.startsWith('string("') && rawValue.endsWith('")')) {
    return rawValue.substring(8, rawValue.length - 2);
  }
  // Strip GP0|#GUID taxonomy prefix — show just the label portion
  if (rawValue.indexOf('GP0|#') >= 0) {
    const parts = rawValue.split('|');
    // The label is typically after the GUID portion
    const lastPart = parts[parts.length - 1];
    return lastPart || rawValue;
  }
  // Strip surrounding quotes
  if (rawValue.charAt(0) === '"' && rawValue.charAt(rawValue.length - 1) === '"') {
    return rawValue.substring(1, rawValue.length - 1);
  }
  // Truncate long values
  if (rawValue.length > 40) {
    return rawValue.substring(0, 37) + '...';
  }
  return rawValue;
}

/**
 * ActiveFilterPillBar — horizontal strip of dismissible filter pills
 * displayed above search results. Multi-value filters are combined
 * into a single pill with comma-separated values.
 */
const ActiveFilterPillBar: React.FC<IActiveFilterPillBarProps> = function ActiveFilterPillBar(props) {
  const { activeFilters, filterConfig, onRemoveFilter, onClearAll, onReset } = props;
  const [displayMap, setDisplayMap] = React.useState<Map<string, string>>(new Map());
  const displayMapRef = React.useRef<Map<string, string>>(displayMap);
  displayMapRef.current = displayMap;

  // Serialize filterConfig IDs to a stable string for comparison.
  // This avoids infinite re-renders when filterConfig is a new array reference
  // each render but contains the same configs.
  const filterConfigKey = React.useMemo(function (): string {
    if (!filterConfig || filterConfig.length === 0) { return ''; }
    const parts: string[] = [];
    for (let i = 0; i < filterConfig.length; i++) {
      parts.push(filterConfig[i].id + ':' + filterConfig[i].filterType);
    }
    return parts.join('|');
  }, [filterConfig]);

  // Clear the display cache when filterConfig actually changes (e.g., filter type swap)
  // so that values get re-resolved with the new formatter.
  React.useEffect(function (): void {
    if (displayMapRef.current.size > 0) {
      setDisplayMap(new Map());
      displayMapRef.current = new Map();
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filterConfigKey]);

  // Resolve display values for active filters via formatters
  // eslint-disable-next-line react-hooks/exhaustive-deps -- displayMapRef.current used via ref to avoid infinite loop
  React.useEffect(() => {
    if (activeFilters.length === 0) {
      return;
    }

    let cancelled = false;

    async function resolveValues(): Promise<void> {
      const current = displayMapRef.current;
      const pending = new Map(current);
      let changed = false;
      for (let i = 0; i < activeFilters.length; i++) {
        const filter = activeFilters[i];
        const key = filter.filterName + '|' + filter.value;
        if (pending.has(key)) {
          continue;
        }
        const config = getFilterConfig(filter.filterName, filterConfig);
        const formatter = getFilterValueFormatter(config.filterType);
        try {
          const formatted = await Promise.resolve(
            formatter.formatForDisplay(filter.value, config)
          );
          pending.set(key, formatted || filter.value);
          changed = true;
        } catch {
          pending.set(key, filter.value);
          changed = true;
        }
      }
      if (!cancelled && changed) {
        setDisplayMap(new Map(pending));
      }
    }

    void resolveValues();

    return () => {
      cancelled = true;
    };
  }, [activeFilters, filterConfig]);

  if (activeFilters.length === 0) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null;
  }

  function getDisplayValue(filterName: string, rawValue: string, displayValue?: string): string {
    if (displayValue) {
      return displayValue;
    }
    const key = filterName + '|' + rawValue;
    return displayMap.get(key) || formatValueForDisplay(rawValue);
  }

  const grouped = groupFiltersByName(activeFilters);
  const pillEntries: Array<{ filterName: string; displayName: string; values: Array<{ value: string; displayValue?: string }> }> = [];

  grouped.forEach(function (values: Array<{ value: string; displayValue?: string }>, filterName: string): void {
    pillEntries.push({
      filterName,
      displayName: getDisplayName(filterName, filterConfig),
      values,
    });
  });

  return (
    <div className={styles.activeFilterPillBar} role="list" aria-label="Active filters" aria-live="polite">
      {pillEntries.map(function (entry): React.ReactElement {
        const displayValues = entry.values.map(function (v): string {
          return getDisplayValue(entry.filterName, v.value, v.displayValue);
        }).join(', ');
        return (
          <div
            key={entry.filterName}
            className={styles.activeFilterPill}
            role="listitem"
          >
            <span
              className={styles.activeFilterPillLabel}
              title={entry.displayName + ': ' + displayValues}
            >
              <strong>{entry.displayName}</strong>: {displayValues}
            </span>
            <button
              className={styles.activeFilterPillRemove}
              onClick={function (): void { onRemoveFilter(entry.filterName); }}
              aria-label={'Remove filter ' + entry.displayName}
              type="button"
            >
              <Icon iconName="Cancel" style={{ fontSize: 10 }} />
            </button>
          </div>
        );
      })}

      {activeFilters.length > 1 && (
        <button
          className={styles.activeFilterClearAll}
          onClick={onClearAll}
          type="button"
          aria-label="Clear all filters"
        >
          Clear all
        </button>
      )}

      <button
        className={styles.activeFilterReset}
        onClick={onReset}
        type="button"
        aria-label="Reset search"
      >
        <Icon iconName="RevToggleKey" style={{ fontSize: 12 }} />
        Reset
      </button>
    </div>
  );
};

export default ActiveFilterPillBar;
