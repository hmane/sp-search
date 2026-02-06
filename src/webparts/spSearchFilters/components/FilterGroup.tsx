import * as React from 'react';
import { Card, Header, Content } from 'spfx-toolkit/lib/components/Card/components';
import styles from './SpSearchFilters.module.scss';
import CheckboxFilter from './CheckboxFilter';
import DateRangeFilter from './DateRangeFilter';
import type {
  IRefiner,
  IActiveFilter,
  IFilterConfig
} from '@interfaces/index';

export interface IFilterGroupProps {
  refiner: IRefiner;
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
}

/**
 * Count the number of active filter selections for a specific filter name.
 */
function countActiveForFilter(filterName: string, activeFilters: IActiveFilter[]): number {
  let count: number = 0;
  for (let i: number = 0; i < activeFilters.length; i++) {
    if (activeFilters[i].filterName === filterName) {
      count++;
    }
  }
  return count;
}

const FilterGroup: React.FC<IFilterGroupProps> = (props: IFilterGroupProps): React.ReactElement => {
  const { refiner, config, activeFilters, onToggleRefiner } = props;

  const displayName: string = config ? config.displayName : refiner.filterName;
  const defaultExpanded: boolean = config ? config.defaultExpanded : true;
  const activeCount: number = countActiveForFilter(refiner.filterName, activeFilters);
  const cardId: string = 'filter-' + refiner.filterName;

  return (
    <Card
      id={cardId}
      defaultExpanded={defaultExpanded}
      allowExpand={true}
      allowMaximize={false}
      elevation={1}
      className={styles.filterGroupCard}
      headerSize="compact"
      animation={{ duration: 250, easing: 'ease-in-out' }}
    >
      <Header
        size="compact"
        hideMaximizeButton={true}
        clickable={true}
      >
        <div className={styles.filterGroupHeader}>
          <span className={styles.filterGroupTitle}>{displayName}</span>
          {activeCount > 0 && (
            <span className={styles.activeCount}>({activeCount})</span>
          )}
        </div>
      </Header>
      <Content padding="compact">
        {config && config.filterType === 'daterange' ? (
          <DateRangeFilter
            filterName={refiner.filterName}
            values={refiner.values}
            config={config}
            activeFilters={activeFilters}
            onToggleRefiner={onToggleRefiner}
          />
        ) : (
          <CheckboxFilter
            filterName={refiner.filterName}
            values={refiner.values}
            config={config}
            activeFilters={activeFilters}
            onToggleRefiner={onToggleRefiner}
          />
        )}
      </Content>
    </Card>
  );
};

export default FilterGroup;
