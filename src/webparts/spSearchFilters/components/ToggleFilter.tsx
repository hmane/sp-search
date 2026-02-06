import * as React from 'react';
import { Toggle } from '@fluentui/react/lib/Toggle';
import styles from './SpSearchFilters.module.scss';
import type { IActiveFilter, IFilterConfig, IRefinerValue } from '@interfaces/index';

export interface IToggleFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
}

const ToggleFilter: React.FC<IToggleFilterProps> = (props: IToggleFilterProps): React.ReactElement => {
  const { filterName, config, activeFilters, onToggleRefiner } = props;

  const operator: 'AND' | 'OR' = config ? config.operator : 'OR';
  const falseLabel = config?.falseLabel || 'No';

  const currentActive: IActiveFilter | undefined = React.useMemo(() => {
    for (let i = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        return activeFilters[i];
      }
    }
    return undefined;
  }, [activeFilters, filterName]);

  const currentValue = currentActive ? currentActive.value : undefined;
  const isYes = currentValue === '1' || currentValue === 'true';
  const isNo = currentValue === '0' || currentValue === 'false';

  function clearFilter(): void {
    if (currentActive) {
      onToggleRefiner(currentActive);
    }
  }

  function setFilterValue(value: string): void {
    if (currentActive && currentActive.value !== value) {
      onToggleRefiner(currentActive);
    }
    const next: IActiveFilter = {
      filterName,
      value,
      operator,
    };
    onToggleRefiner(next);
  }

  function handleToggleChange(_: React.MouseEvent<HTMLElement>, checked?: boolean): void {
    if (checked) {
      setFilterValue('1');
    } else {
      clearFilter();
    }
  }

  function handleNoClick(): void {
    setFilterValue('0');
  }

  function handleAllClick(): void {
    clearFilter();
  }

  return (
    <div className={styles.toggleFilterContainer}>
      <Toggle
        label={config ? config.displayName : filterName}
        checked={!!isYes}
        onChange={handleToggleChange}
        inlineLabel={true}
      />
      <div className={styles.toggleStateRow}>
        <button
          type="button"
          className={isNo ? styles.toggleStateButtonActive : styles.toggleStateButton}
          onClick={handleNoClick}
        >
          {falseLabel}
        </button>
        <button
          type="button"
          className={!isYes && !isNo ? styles.toggleStateButtonActive : styles.toggleStateButton}
          onClick={handleAllClick}
        >
          All
        </button>
      </div>
    </div>
  );
};

export default ToggleFilter;
