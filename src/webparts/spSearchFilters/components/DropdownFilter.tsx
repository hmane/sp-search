import * as React from 'react';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import type { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import styles from './SpSearchFilters.module.scss';
import { getSelectedRefinerTokens } from './filterSelectionUtils';
import type {
  IRefinerValue,
  IActiveFilter,
  IFilterConfig,
  SortBy
} from '@interfaces/index';

export interface IDropdownFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
  /**
   * Batched callback (Task 1 foundation). Components migrating in Tasks 2-5
   * will switch from per-delta `onToggleRefiner` to a single batched call here.
   */
  onReplaceRefinerValues?: (payload: { filterName: string; values: IActiveFilter[] }) => void;
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

const DropdownFilter: React.FC<IDropdownFilterProps> = (props: IDropdownFilterProps): React.ReactElement => {
  const { filterName, values, config, activeFilters, onToggleRefiner } = props;

  const showCount: boolean = config ? config.showCount : true;
  const sortBy: SortBy = config ? config.sortBy : 'count';
  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';
  const allowMultiple: boolean = config?.multiValues === true;

  const [isExpanded, setIsExpanded] = React.useState<boolean>(false);

  const sortedValues: IRefinerValue[] = React.useMemo(function (): IRefinerValue[] {
    const sorted = values.slice();
    if (sortBy === 'alphabetical') {
      sorted.sort(compareAlphabetical);
    } else {
      sorted.sort(compareByCount);
    }
    return sorted;
  }, [sortBy, values]);

  const selectedValues = React.useMemo(function (): string[] {
    return getSelectedRefinerTokens(filterName, values, activeFilters);
  }, [activeFilters, filterName, values]);

  const maxVisible: number = config && config.maxValues > 0 ? config.maxValues : 10;
  const hasMore: boolean = sortedValues.length > maxVisible;

  const options = React.useMemo(function (): IDropdownOption[] {
    const selectedSet = new Set(selectedValues);
    const limit = isExpanded ? sortedValues.length : maxVisible;
    const limitedValues = sortedValues.filter(function (value: IRefinerValue, index: number): boolean {
      return index < limit || selectedSet.has(value.value);
    });
    return limitedValues.map(function (value: IRefinerValue): IDropdownOption {
      const label = value.name || value.value;
      return {
        key: value.value,
        text: showCount ? label + ' (' + String(value.count) + ')' : label
      };
    });
  }, [isExpanded, maxVisible, selectedValues, showCount, sortedValues]);

  const displayNameMap = React.useMemo(function (): Map<string, string> {
    const map = new Map<string, string>();
    for (let i = 0; i < sortedValues.length; i++) {
      map.set(sortedValues[i].value, sortedValues[i].name || sortedValues[i].value);
    }
    return map;
  }, [sortedValues]);

  function syncToSelection(nextSelectedValues: string[]): void {
    for (let i = 0; i < nextSelectedValues.length; i++) {
      if (selectedValues.indexOf(nextSelectedValues[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: nextSelectedValues[i],
          displayValue: displayNameMap.get(nextSelectedValues[i]) || undefined,
          operator
        });
      }
    }

    for (let i = 0; i < selectedValues.length; i++) {
      if (nextSelectedValues.indexOf(selectedValues[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: selectedValues[i],
          displayValue: displayNameMap.get(selectedValues[i]) || undefined,
          operator
        });
      }
    }
  }

  function handleSingleChange(
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    const nextSelectedValues = option ? [String(option.key)] : [];
    syncToSelection(nextSelectedValues);
  }

  function handleMultiChange(
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    if (!option) {
      return;
    }
    const optionKey = String(option.key);
    const nextSelectedValues = option.selected
      ? selectedValues.concat(optionKey)
      : selectedValues.filter(function (value: string): boolean { return value !== optionKey; });
    syncToSelection(nextSelectedValues);
  }

  return (
    <div className={styles.dropdownFilterContainer}>
      <Dropdown
        placeholder={allowMultiple ? 'Select values...' : 'Select value...'}
        options={options}
        selectedKey={allowMultiple ? undefined : (selectedValues[0] || null)}
        selectedKeys={allowMultiple ? selectedValues : undefined}
        multiSelect={allowMultiple}
        onChange={allowMultiple ? handleMultiChange : handleSingleChange}
        className={styles.dropdownFilter}
      />
      {options.length === 0 && (
        <div className={styles.noResults}>No matching values</div>
      )}
      {hasMore && (
        <button
          type="button"
          className={styles.showMoreBtn}
          onClick={function (): void { setIsExpanded(!isExpanded); }}
          aria-expanded={isExpanded}
        >
          {isExpanded ? 'Show less' : 'Show more (' + (sortedValues.length - maxVisible) + ')'}
        </button>
      )}
    </div>
  );
};

export default DropdownFilter;
