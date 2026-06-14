import * as React from 'react';
import { Card, Header, Content } from 'spfx-toolkit/lib/components/Card';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './SpSearchFilters.module.scss';
import TextFilter from './TextFilter';
import ToggleFilter from './ToggleFilter';
import type {
  IRefiner,
  IActiveFilter,
  IFilterConfig,
  IReplaceRefinerValuesPayload
} from '@interfaces/index';

// Lazy-load heavy filter components (DevExtreme DateBox, RangeSlider, TagBox, TreeView; PnP FileTypeIcon, PeoplePicker)
const LazyCheckboxFilter = React.lazy(function () { return import(/* webpackChunkName: 'CheckboxFilter' */ './CheckboxFilter'); });
const LazyDateRangeFilter = React.lazy(function () { return import(/* webpackChunkName: 'DateRangeFilter' */ './DateRangeFilter'); });
const LazySliderFilter = React.lazy(function () { return import(/* webpackChunkName: 'SliderFilter' */ './SliderFilter'); });
const LazyTagBoxFilter = React.lazy(function () { return import(/* webpackChunkName: 'TagBoxFilter' */ './TagBoxFilter'); });
const LazyTaxonomyTreeFilter = React.lazy(function () { return import(/* webpackChunkName: 'TaxonomyTreeFilter' */ './TaxonomyTreeFilter'); });
const LazyPeoplePickerFilter = React.lazy(function () { return import(/* webpackChunkName: 'PeoplePickerFilter' */ './PeoplePickerFilter'); });
const LazyDropdownFilter = React.lazy(function () { return import(/* webpackChunkName: 'DropdownFilter' */ './DropdownFilter'); });

export interface IFilterGroupProps {
  refiner: IRefiner;
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
  /**
   * Multi-value batched callback. Filter components that mutate more than one
   * value per user interaction (TagBox, People, Taxonomy, multi-Dropdown) call
   * this with the full intended selection for their `filterName`, avoiding the
   * stale-closure clobber that affects per-delta `onToggleRefiner` calls.
   * Optional during the Task 1 foundation; components migrate in Tasks 2-5.
   */
  onReplaceRefinerValues?: (payload: IReplaceRefinerValuesPayload) => void;
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
  onToggleRefiner: (filter: IActiveFilter) => void,
  onReplaceRefinerValues: ((payload: IReplaceRefinerValuesPayload) => void) | undefined
): React.ReactElement {
  const filterType: string = config ? config.filterType : 'checkbox';
  const commonProps = {
    filterName: refiner.filterName,
    values: refiner.values,
    config: config,
    activeFilters: activeFilters,
    onToggleRefiner: onToggleRefiner,
    onReplaceRefinerValues: onReplaceRefinerValues
  };

  switch (filterType) {
    case 'dropdown':
      return React.createElement(
        React.Suspense,
        { fallback: LazyFallback },
        React.createElement(LazyDropdownFilter, commonProps)
      );
    case 'daterange':
      return React.createElement(
        React.Suspense,
        { fallback: LazyFallback },
        React.createElement(LazyDateRangeFilter, commonProps)
      );
    case 'text':
      return React.createElement(TextFilter, commonProps);
    case 'toggle':
      return React.createElement(ToggleFilter, commonProps);
    case 'tagbox':
      return React.createElement(
        React.Suspense,
        { fallback: LazyFallback },
        React.createElement(LazyTagBoxFilter, commonProps)
      );
    case 'slider':
      return React.createElement(
        React.Suspense,
        { fallback: LazyFallback },
        React.createElement(LazySliderFilter, commonProps)
      );
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
      return React.createElement(
        React.Suspense,
        { fallback: LazyFallback },
        React.createElement(LazyCheckboxFilter, commonProps)
      );
  }
}

const FilterGroup: React.FC<IFilterGroupProps> = (props: IFilterGroupProps): React.ReactElement => {
  const { refiner, config, activeFilters, onToggleRefiner, onReplaceRefinerValues } = props;

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
        {renderFilterComponent(refiner, config, activeFilters, onToggleRefiner, onReplaceRefinerValues)}
      </Content>
    </Card>
  );
};

export default FilterGroup;
