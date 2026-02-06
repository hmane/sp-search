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
  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';
  const sortBy: SortBy = config ? config.sortBy : 'count';

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

  const [currentSelection, setCurrentSelection] = React.useState<string[]>(selectedValues);

  React.useEffect(() => {
    setCurrentSelection(selectedValues);
  }, [selectedValues]);

  function handleValueChanged(e: { value?: string[] }): void {
    const nextValues: string[] = Array.isArray(e.value) ? e.value : [];
    const prevValues = currentSelection;

    for (let i = 0; i < nextValues.length; i++) {
      if (prevValues.indexOf(nextValues[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: nextValues[i],
          operator,
        });
      }
    }

    for (let i = 0; i < prevValues.length; i++) {
      if (nextValues.indexOf(prevValues[i]) < 0) {
        onToggleRefiner({
          filterName,
          value: prevValues[i],
          operator,
        });
      }
    }

    setCurrentSelection(nextValues);
  }

  return (
    <div className={styles.tagBoxFilterContainer}>
      <TagBox
        dataSource={items}
        valueExpr="value"
        displayExpr="displayName"
        showSelectionControls={true}
        searchEnabled={true}
        multiline={true}
        hideSelectedItems={false}
        value={currentSelection}
        onValueChanged={handleValueChanged}
        placeholder="Select values"
        height={120}
      />
    </div>
  );
};

export default TagBoxFilter;
