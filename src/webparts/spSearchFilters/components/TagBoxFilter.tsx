import * as React from 'react';
import { TagBox } from 'devextreme-react/tag-box';
import styles from './SpSearchFilters.module.scss';
import type {
  IRefinerValue,
  IActiveFilter,
  IFilterConfig,
  SortBy
} from '@interfaces/index';

export interface ITagBoxFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
}

function compareByCount(a: IRefinerValue, b: IRefinerValue): number {
  return b.count - a.count;
}

function compareAlphabetical(a: IRefinerValue, b: IRefinerValue): number {
  const nameA: string = (a.name || '').toLowerCase();
  const nameB: string = (b.name || '').toLowerCase();
  if (nameA < nameB) { return -1; }
  if (nameA > nameB) { return 1; }
  return 0;
}

const TagBoxFilter: React.FC<ITagBoxFilterProps> = (props: ITagBoxFilterProps): React.ReactElement => {
  const { filterName, values, config, activeFilters, onToggleRefiner } = props;

  const showCount: boolean = config ? config.showCount : true;
  const configSortBy: SortBy = config ? config.sortBy : 'count';
  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';

  const [sortBy, setSortBy] = React.useState<SortBy>(configSortBy);

  const sortedValues: IRefinerValue[] = React.useMemo(function (): IRefinerValue[] {
    const sorted: IRefinerValue[] = values.slice();
    if (sortBy === 'alphabetical') {
      sorted.sort(compareAlphabetical);
    } else {
      sorted.sort(compareByCount);
    }
    return sorted;
  }, [values, sortBy]);

  const items = React.useMemo(function (): Array<{ value: string; displayName: string }> {
    return sortedValues.map((value) => {
      const name = value.name || value.value;
      return {
        value: value.value,
        displayName: showCount ? name + ' (' + String(value.count) + ')' : name,
      };
    });
  }, [sortedValues, showCount]);

  const selectedValues = React.useMemo(function (): string[] {
    const selected: string[] = [];
    for (let i = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        selected.push(activeFilters[i].value);
      }
    }
    return selected;
  }, [activeFilters, filterName]);

  // Guard against re-entrant onValueChanged calls from programmatic value updates
  const isUpdatingRef = React.useRef<boolean>(false);

  function handleValueChanged(e: { value?: string[] }): void {
    if (isUpdatingRef.current) {
      return;
    }
    isUpdatingRef.current = true;

    const nextValues: string[] = Array.isArray(e.value) ? e.value : [];

    // Toggle added values
    for (let i = 0; i < nextValues.length; i++) {
      if (selectedValues.indexOf(nextValues[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: nextValues[i],
          operator,
        });
      }
    }

    // Toggle removed values
    for (let i = 0; i < selectedValues.length; i++) {
      if (nextValues.indexOf(selectedValues[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: selectedValues[i],
          operator,
        });
      }
    }

    // Release guard on next tick so React can re-render with store values
    setTimeout(function (): void { isUpdatingRef.current = false; }, 0);
  }

  function handleSortByCount(): void {
    setSortBy('count');
  }

  function handleSortAlphabetical(): void {
    setSortBy('alphabetical');
  }

  return (
    <div className={styles.tagBoxFilterContainer}>
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
      <TagBox
        dataSource={items}
        valueExpr="value"
        displayExpr="displayName"
        showSelectionControls={true}
        showDropDownButton={true}
        searchEnabled={true}
        multiline={true}
        hideSelectedItems={false}
        value={selectedValues}
        onValueChanged={handleValueChanged}
        placeholder="Select values..."
        maxDisplayedTags={5}
        showMultiTagOnly={false}
      />
    </div>
  );
};

export default TagBoxFilter;
