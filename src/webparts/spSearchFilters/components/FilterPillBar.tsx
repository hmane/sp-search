import * as React from 'react';
import styles from './SpSearchFilters.module.scss';
import type { IActiveFilter, IFilterConfig } from '@interfaces/index';

export interface IFilterPillBarProps {
  activeFilters: IActiveFilter[];
  filterConfig: IFilterConfig[];
  onRemoveFilter: (filterName: string, value: string) => void;
  onClearAll: () => void;
  showClearAll: boolean;
}

/**
 * Resolves a filter name to its display name using filterConfig.
 * Falls back to the raw filterName if no config match is found.
 */
function getDisplayName(filterName: string, filterConfig: IFilterConfig[]): string {
  for (let i: number = 0; i < filterConfig.length; i++) {
    if (filterConfig[i].managedProperty === filterName) {
      return filterConfig[i].displayName;
    }
  }
  return filterName;
}

const FilterPillBar: React.FC<IFilterPillBarProps> = (props: IFilterPillBarProps): React.ReactElement => {
  const { activeFilters, filterConfig, onRemoveFilter, onClearAll, showClearAll } = props;

  if (activeFilters.length === 0) {
    return <></>;
  }

  return (
    <div className={styles.pillBar} role="list" aria-label="Active filters">
      {activeFilters.map(function (filter: IActiveFilter, index: number): React.ReactElement {
        const displayName: string = getDisplayName(filter.filterName, filterConfig);
        const key: string = filter.filterName + '_' + filter.value + '_' + index;
        return (
          <div key={key} className={styles.pill} role="listitem">
            <span className={styles.pillLabel} title={displayName + ': ' + filter.value}>
              {displayName}: {filter.value}
            </span>
            <button
              className={styles.pillRemove}
              onClick={function (): void { onRemoveFilter(filter.filterName, filter.value); }}
              aria-label={'Remove filter ' + displayName + ' ' + filter.value}
              type="button"
            >
              &#x2715;
            </button>
          </div>
        );
      })}
      {showClearAll && activeFilters.length > 1 && (
        <button
          className={styles.clearAllLink}
          onClick={onClearAll}
          type="button"
          aria-label="Clear all filters"
        >
          Clear all
        </button>
      )}
    </div>
  );
};

export default FilterPillBar;
