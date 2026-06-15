import * as React from 'react';
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
  const trueLabel = config?.trueLabel || 'Yes';
  const falseLabel = config?.falseLabel || 'No';
  const invertBoolean = config?.invertBoolean === true;

  const currentActive: IActiveFilter | undefined = React.useMemo(() => {
    for (let i = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        return activeFilters[i];
      }
    }
    return undefined;
  }, [activeFilters, filterName]);

  const currentValue = currentActive ? currentActive.value : undefined;
  const rawIsTrue = currentValue === '1' || currentValue === 'true';
  const rawIsFalse = currentValue === '0' || currentValue === 'false';
  const isYes = invertBoolean ? rawIsFalse : rawIsTrue;
  const isNo = invertBoolean ? rawIsTrue : rawIsFalse;

  function clearFilter(): void {
    if (currentActive) {
      onToggleRefiner(currentActive);
    }
  }

  function setFilterValue(value: string): void {
    if (currentActive && currentActive.value === value) {
      return;
    }
    if (currentActive && currentActive.value !== value) {
      onToggleRefiner(currentActive);
    }
    const next: IActiveFilter = {
      filterName,
      value,
      displayValue: (value === '1' || value === 'true')
        ? (invertBoolean ? falseLabel : trueLabel)
        : (invertBoolean ? trueLabel : falseLabel),
      operator,
    };
    onToggleRefiner(next);
  }

  function handleYesClick(): void {
    setFilterValue(invertBoolean ? '0' : '1');
  }

  function handleNoClick(): void {
    setFilterValue(invertBoolean ? '1' : '0');
  }

  function handleAllClick(): void {
    clearFilter();
  }

  return (
    <div className={styles.toggleFilterContainer} role="group" aria-label={config ? config.displayName : filterName}>
      <div className={styles.toggleStateRow}>
        <button
          type="button"
          className={!isYes && !isNo ? styles.toggleStateButtonActive : styles.toggleStateButton}
          onClick={handleAllClick}
          aria-pressed={!isYes && !isNo}
        >
          All
        </button>
        <button
          type="button"
          className={isYes ? styles.toggleStateButtonActive : styles.toggleStateButton}
          onClick={handleYesClick}
          aria-pressed={isYes}
        >
          {trueLabel}
        </button>
        <button
          type="button"
          className={isNo ? styles.toggleStateButtonActive : styles.toggleStateButton}
          onClick={handleNoClick}
          aria-pressed={isNo}
        >
          {falseLabel}
        </button>
      </div>
    </div>
  );
};

export default ToggleFilter;
