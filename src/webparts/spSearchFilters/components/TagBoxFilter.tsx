import * as React from 'react';
import { TagBox } from 'devextreme-react/tag-box';
import styles from './SpSearchFilters.module.scss';
import { getSelectedRefinerTokens } from './filterSelectionUtils';
import type {
  IRefinerValue,
  IActiveFilter,
  IFilterConfig,
  IReplaceRefinerValuesPayload,
  SortBy
} from '@interfaces/index';

export interface ITagBoxFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
  /**
   * Batched callback (Task 1 foundation). Components migrating in Tasks 2-5
   * will switch from per-delta `onToggleRefiner` to a single batched call here.
   */
  onReplaceRefinerValues?: (payload: IReplaceRefinerValuesPayload) => void;
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

/**
 * Pure helper that builds the batched payload for `onReplaceRefinerValues`
 * from the editor's full intended selection. Extracted so the multi-select
 * cascade can be tested without needing to drive DevExtreme TagBox under
 * jsdom (its onValueChanged is hard to reach via @testing-library).
 *
 * Order of `nextSelectedTokens` is preserved (matches what the user sees
 * in the TagBox). Display labels are looked up from the current refiner
 * bucket; tokens not present in the bucket get `displayValue: undefined`
 * (the URL/pill formatter will fall back to the raw token).
 */
export function buildTagBoxBatchPayload(input: {
  filterName: string;
  nextSelectedTokens: string[];
  refinerValues: IRefinerValue[];
  operator: 'AND' | 'OR';
}): IReplaceRefinerValuesPayload {
  const labelMap: Map<string, string> = new Map<string, string>();
  for (let i = 0; i < input.refinerValues.length; i++) {
    const rv = input.refinerValues[i];
    labelMap.set(rv.value, rv.name || rv.value);
  }

  const values: IActiveFilter[] = [];
  for (let i = 0; i < input.nextSelectedTokens.length; i++) {
    const token: string = input.nextSelectedTokens[i];
    values.push({
      filterName: input.filterName,
      value: token,
      displayValue: labelMap.get(token),
      operator: input.operator,
    });
  }

  return { filterName: input.filterName, values: values };
}

const TagBoxFilter: React.FC<ITagBoxFilterProps> = (props: ITagBoxFilterProps): React.ReactElement => {
  const { filterName, values, config, activeFilters, onToggleRefiner, onReplaceRefinerValues } = props;

  const showCount: boolean = config ? config.showCount : true;
  const configSortBy: SortBy = config ? config.sortBy : 'count';
  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';
  const allowMultiple: boolean = config?.multiValues !== false;

  const [sortBy, setSortBy] = React.useState<SortBy>(configSortBy);
  const [isExpanded, setIsExpanded] = React.useState<boolean>(false);

  const sortedValues: IRefinerValue[] = React.useMemo(function (): IRefinerValue[] {
    const sorted: IRefinerValue[] = values.slice();
    if (sortBy === 'alphabetical') {
      sorted.sort(compareAlphabetical);
    } else {
      sorted.sort(compareByCount);
    }
    return sorted;
  }, [values, sortBy]);

  const selectedValues = React.useMemo(function (): string[] {
    return getSelectedRefinerTokens(filterName, values, activeFilters);
  }, [activeFilters, filterName, values]);

  const [editorValues, setEditorValues] = React.useState<string[]>(selectedValues);

  React.useEffect(function (): void {
    setEditorValues(selectedValues);
  }, [selectedValues]);

  const maxVisible: number = config && config.maxValues > 0 ? config.maxValues : 10;
  const hasMore: boolean = sortedValues.length > maxVisible;

  const items = React.useMemo(function (): Array<{ value: string; displayName: string }> {
    const selectedSet = new Set(selectedValues);
    const limit = isExpanded ? sortedValues.length : maxVisible;
    const limitedValues = sortedValues.filter(function (value: IRefinerValue, index: number): boolean {
      return index < limit || selectedSet.has(value.value);
    });

    return limitedValues.map((value) => {
      const name = value.name || value.value;
      return {
        value: value.value,
        displayName: showCount ? name + ' (' + String(value.count) + ')' : name,
      };
    });
  }, [isExpanded, maxVisible, selectedValues, showCount, sortedValues]);

  // Guard against re-entrant onValueChanged calls from programmatic value updates
  const isUpdatingRef = React.useRef<boolean>(false);

  // Build a value→displayName lookup for setting displayValue on active filters
  const displayNameMap = React.useMemo(function (): Map<string, string> {
    const map = new Map<string, string>();
    for (let i = 0; i < sortedValues.length; i++) {
      const name = sortedValues[i].name || sortedValues[i].value;
      map.set(sortedValues[i].value, name);
    }
    return map;
  }, [sortedValues]);

  function handleValueChanged(e: { value?: string[] }): void {
    if (isUpdatingRef.current) {
      return;
    }
    isUpdatingRef.current = true;

    let nextValues: string[] = Array.isArray(e.value) ? e.value : [];
    if (!allowMultiple && nextValues.length > 1) {
      nextValues = [nextValues[nextValues.length - 1]];
    }
    setEditorValues(nextValues);

    if (onReplaceRefinerValues) {
      // Batched path — one call carrying the FULL intended selection.
      // The per-delta loop below would otherwise capture a stale React
      // closure over `activeFilters` and silently drop all but the last
      // delta on multi-select (paste-multi, programmatic selection).
      const payload = buildTagBoxBatchPayload({
        filterName,
        nextSelectedTokens: nextValues,
        refinerValues: sortedValues,
        operator,
      });
      onReplaceRefinerValues(payload);
    } else {
      // Fallback for callers that haven't migrated. Kept so the component
      // remains usable standalone, but the in-tree parent always provides
      // onReplaceRefinerValues.

      // Toggle added values
      for (let i = 0; i < nextValues.length; i++) {
        if (selectedValues.indexOf(nextValues[i]) < 0) {
          onToggleRefiner({
            filterName,
            value: nextValues[i],
            displayValue: displayNameMap.get(nextValues[i]) || undefined,
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
            displayValue: displayNameMap.get(selectedValues[i]) || undefined,
            operator,
          });
        }
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
        value={editorValues}
        onValueChanged={handleValueChanged}
        applyValueMode="useButtons"
        placeholder="Select values..."
        maxDisplayedTags={5}
        showMultiTagOnly={false}
      />
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

export default TagBoxFilter;
