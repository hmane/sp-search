import * as React from 'react';
import { Card, Header, Content } from 'spfx-toolkit/lib/components/Card/components';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './SpSearchFilters.module.scss';
import CheckboxFilter from './CheckboxFilter';
import DateRangeFilter from './DateRangeFilter';
import ToggleFilter from './ToggleFilter';
import SliderFilter from './SliderFilter';
import type {
  IRefiner,
  IActiveFilter,
  IFilterConfig
} from '@interfaces/index';

// Lazy-load heavy filter components (DevExtreme TagBox, DevExtreme TreeView, PnP PeoplePicker)
const LazyTagBoxFilter = React.lazy(function () { return import('./TagBoxFilter'); });
const LazyTaxonomyTreeFilter = React.lazy(function () { return import('./TaxonomyTreeFilter'); });
const LazyPeoplePickerFilter = React.lazy(function () { return import('./PeoplePickerFilter'); });

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

/** Fallback spinner shown while lazy filter components load. */
const LazyFallback: React.ReactElement = React.createElement(
  Spinner,
  { size: SpinnerSize.small, label: 'Loading filter...' }
);

/**
 * Render the appropriate filter component based on filterType config.
 */
function renderFilterComponent(
  refiner: IRefiner,
  config: IFilterConfig | undefined,
  activeFilters: IActiveFilter[],
  onToggleRefiner: (filter: IActiveFilter) => void
): React.ReactElement {
  const filterType: string = config ? config.filterType : 'checkbox';
  const commonProps = {
    filterName: refiner.filterName,
    values: refiner.values,
    config: config,
    activeFilters: activeFilters,
    onToggleRefiner: onToggleRefiner
  };

  switch (filterType) {
    case 'daterange':
      return React.createElement(DateRangeFilter, commonProps);
    case 'toggle':
      return React.createElement(ToggleFilter, commonProps);
    case 'tagbox':
      return React.createElement(
        React.Suspense,
        { fallback: LazyFallback },
        React.createElement(LazyTagBoxFilter, commonProps)
      );
    case 'slider':
      return React.createElement(SliderFilter, commonProps);
    case 'taxonomy':
      return React.createElement(
        React.Suspense,
        { fallback: LazyFallback },
        React.createElement(LazyTaxonomyTreeFilter, commonProps)
      );
    case 'people':
      return React.createElement(
        React.Suspense,
        { fallback: LazyFallback },
        React.createElement(LazyPeoplePickerFilter, commonProps)
      );
    case 'checkbox':
    default:
      return React.createElement(CheckboxFilter, commonProps);
  }
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
        {renderFilterComponent(refiner, config, activeFilters, onToggleRefiner)}
      </Content>
    </Card>
  );
};

export default FilterGroup;
