import * as React from 'react';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SpSearchFilters.module.scss';
import type {
  IRefinerValue,
  IActiveFilter,
  IFilterConfig,
  SortBy
} from '@interfaces/index';

export interface ICheckboxFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
}

const DEFAULT_VISIBLE: number = 10;

/**
 * Compares two refiner values for sorting by count (descending).
 */
function compareByCount(a: IRefinerValue, b: IRefinerValue): number {
  return b.count - a.count;
}

/**
 * Compares two refiner values for alphabetical sorting (ascending).
 */
function compareAlphabetical(a: IRefinerValue, b: IRefinerValue): number {
  const nameA: string = (a.name || '').toLowerCase();
  const nameB: string = (b.name || '').toLowerCase();
  if (nameA < nameB) { return -1; }
  if (nameA > nameB) { return 1; }
  return 0;
}

const CheckboxFilter: React.FC<ICheckboxFilterProps> = (props: ICheckboxFilterProps): React.ReactElement => {
  const { filterName, values, config, activeFilters, onToggleRefiner } = props;

  const showCount: boolean = config ? config.showCount : true;
  const configSortBy: SortBy = config ? config.sortBy : 'count';
  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';

  const [searchText, setSearchText] = React.useState<string>('');
  const [isExpanded, setIsExpanded] = React.useState<boolean>(false);
  const [sortBy, setSortBy] = React.useState<SortBy>(configSortBy);

  /** Determine if a value is currently selected in activeFilters. */
  function isValueSelected(value: string): boolean {
    for (let i: number = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName && activeFilters[i].value === value) {
        return true;
      }
    }
    return false;
  }

  /** Filter values by search text. */
  const filteredValues: IRefinerValue[] = React.useMemo(function (): IRefinerValue[] {
    if (!searchText) {
      return values;
    }
    const lowerSearch: string = searchText.toLowerCase();
    const result: IRefinerValue[] = [];
    for (let i: number = 0; i < values.length; i++) {
      if ((values[i].name || '').toLowerCase().indexOf(lowerSearch) >= 0) {
        result.push(values[i]);
      }
    }
    return result;
  }, [values, searchText]);

  /** Sort the filtered values. */
  const sortedValues: IRefinerValue[] = React.useMemo(function (): IRefinerValue[] {
    const sorted: IRefinerValue[] = filteredValues.slice();
    if (sortBy === 'alphabetical') {
      sorted.sort(compareAlphabetical);
    } else {
      sorted.sort(compareByCount);
    }
    return sorted;
  }, [filteredValues, sortBy]);

  /** Determine visible values based on show more/less. */
  const visibleValues: IRefinerValue[] = React.useMemo(function (): IRefinerValue[] {
    if (isExpanded || sortedValues.length <= DEFAULT_VISIBLE) {
      return sortedValues;
    }
    return sortedValues.slice(0, DEFAULT_VISIBLE);
  }, [sortedValues, isExpanded]);

  const hasMore: boolean = sortedValues.length > DEFAULT_VISIBLE;

  function handleCheckboxChange(value: IRefinerValue): void {
    const filter: IActiveFilter = {
      filterName: filterName,
      value: value.value,
      operator: operator
    };
    onToggleRefiner(filter);
  }

  function handleSearchChange(ev: React.ChangeEvent<HTMLInputElement>): void {
    setSearchText(ev.target.value);
  }

  function handleShowMore(): void {
    setIsExpanded(!isExpanded);
  }

  function handleSortByCount(): void {
    setSortBy('count');
  }

  function handleSortAlphabetical(): void {
    setSortBy('alphabetical');
  }

  return (
    <div>
      {/* Sort controls */}
      <div className={styles.sortControls}>
        <button
          type="button"
          className={sortBy === 'count' ? styles.sortButtonActive : styles.sortButton}
          onClick={handleSortByCount}
          aria-label="Sort by count"
          aria-pressed={sortBy === 'count'}
        >
          By count
        </button>
        <button
          type="button"
          className={sortBy === 'alphabetical' ? styles.sortButtonActive : styles.sortButton}
          onClick={handleSortAlphabetical}
          aria-label="Sort alphabetically"
          aria-pressed={sortBy === 'alphabetical'}
        >
          A-Z
        </button>
      </div>

      {/* Search within filter */}
      {values.length > DEFAULT_VISIBLE && (
        <div className={styles.searchWithinFilter}>
          <Icon iconName="Search" className={styles.searchIcon} />
          <input
            type="text"
            className={styles.searchInput}
            placeholder="Search within this filter"
            value={searchText}
            onChange={handleSearchChange}
            aria-label={'Search within ' + (config ? config.displayName : filterName)}
          />
        </div>
      )}

      {/* Checkbox list */}
      {visibleValues.length === 0 && (
        <div className={styles.noResults}>No matching values</div>
      )}
      <ul className={styles.checkboxList} role="group" aria-label={config ? config.displayName : filterName}>
        {visibleValues.map(function (refinerValue: IRefinerValue): React.ReactElement {
          const checked: boolean = isValueSelected(refinerValue.value);
          return (
            <li key={refinerValue.value} className={styles.checkboxItem}>
              <Checkbox
                className={styles.checkboxLabel}
                label={refinerValue.name || refinerValue.value}
                checked={checked}
                onChange={function (): void { handleCheckboxChange(refinerValue); }}
                ariaLabel={refinerValue.name || refinerValue.value}
              />
              {showCount && (
                <span className={styles.refinerCount}>({refinerValue.count})</span>
              )}
            </li>
          );
        })}
      </ul>

      {/* Show more / Show less */}
      {hasMore && !searchText && (
        <button
          type="button"
          className={styles.showMoreBtn}
          onClick={handleShowMore}
          aria-expanded={isExpanded}
        >
          {isExpanded ? 'Show less' : 'Show more (' + (sortedValues.length - DEFAULT_VISIBLE) + ')'}
        </button>
      )}
    </div>
  );
};

export default CheckboxFilter;
