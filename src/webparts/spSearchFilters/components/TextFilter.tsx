import * as React from 'react';
import styles from './SpSearchFilters.module.scss';
import type { IActiveFilter, IFilterConfig, IRefinerValue } from '@interfaces/index';

export interface ITextFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
}

const TextFilter: React.FC<ITextFilterProps> = (props: ITextFilterProps): React.ReactElement => {
  const { filterName, config, activeFilters, onToggleRefiner } = props;

  const operator: 'AND' | 'OR' = config ? config.operator : 'AND';
  const currentActive = React.useMemo((): IActiveFilter | undefined => {
    for (let i = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        return activeFilters[i];
      }
    }
    return undefined;
  }, [activeFilters, filterName]);

  const [value, setValue] = React.useState<string>(currentActive?.value || '');

  React.useEffect((): void => {
    setValue(currentActive?.value || '');
  }, [currentActive?.value]);

  function applyValue(): void {
    const trimmed = value.trim();
    if (!trimmed) {
      if (currentActive) {
        onToggleRefiner(currentActive);
      }
      return;
    }

    // Re-click with the same value toggles it off in the parent/store layer.
    onToggleRefiner({
      filterName,
      value: trimmed,
      displayValue: trimmed,
      operator
    });
  }

  function clearValue(): void {
    setValue('');
    if (currentActive) {
      onToggleRefiner(currentActive);
    }
  }

  return (
    <div className={styles.textFilterContainer}>
      <div className={styles.textFilterInputRow}>
        <input
          type="text"
          className={styles.searchInput}
          placeholder={'Search ' + (config?.displayName || filterName)}
          value={value}
          onChange={(ev): void => { setValue(ev.target.value); }}
          onKeyDown={(ev): void => {
            if (ev.key === 'Enter') {
              ev.preventDefault();
              applyValue();
            }
          }}
          aria-label={'Search ' + (config?.displayName || filterName)}
        />
        <button
          type="button"
          className={styles.textFilterButton}
          onClick={applyValue}
        >
          Apply
        </button>
      </div>
      {(currentActive || value) && (
        <button
          type="button"
          className={styles.textFilterClear}
          onClick={clearValue}
        >
          Clear
        </button>
      )}
    </div>
  );
};

export default TextFilter;
