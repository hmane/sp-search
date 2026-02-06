import * as React from 'react';
import { RangeSlider, Label, Tooltip } from 'devextreme-react/range-slider';
import styles from './SpSearchFilters.module.scss';
import type { IRefinerValue, IActiveFilter, IFilterConfig } from '@interfaces/index';
import { formatNumericValue, parseRangeToken } from '@store/formatters/FilterValueFormatters';

export interface ISliderFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
}

/**
 * Extract numeric values from refiner buckets.
 */
function extractNumericValues(values: IRefinerValue[]): number[] {
  const nums: number[] = [];
  for (let i: number = 0; i < values.length; i++) {
    const parsed: number = parseFloat(values[i].value);
    if (!isNaN(parsed)) {
      nums.push(parsed);
    }
  }
  return nums;
}

/**
 * SliderFilter — DevExtreme RangeSlider for numeric range refinement.
 *
 * Features:
 * - Configurable min/max derived from refiner values
 * - FQL range(decimal(...), decimal(...)) token generation
 * - File size formatting when managed property contains 'size'
 */
const SliderFilter: React.FC<ISliderFilterProps> = (props: ISliderFilterProps): React.ReactElement => {
  const { filterName, values, config, activeFilters, onToggleRefiner } = props;

  const numericValues: number[] = React.useMemo(function (): number[] {
    return extractNumericValues(values);
  }, [values]);

  const rangeMin: number = React.useMemo(function (): number {
    if (config && typeof config.rangeMin === 'number') {
      return config.rangeMin;
    }
    if (numericValues.length > 0) {
      return Math.min.apply(null, numericValues);
    }
    return 0;
  }, [config, numericValues]);

  const rangeMax: number = React.useMemo(function (): number {
    if (config && typeof config.rangeMax === 'number') {
      return config.rangeMax;
    }
    if (numericValues.length > 0) {
      return Math.max.apply(null, numericValues);
    }
    return 100;
  }, [config, numericValues]);

  // Compute step size based on range
  const step: number = React.useMemo(function (): number {
    if (config && typeof config.rangeStep === 'number') {
      return config.rangeStep;
    }
    const range: number = rangeMax - rangeMin;
    if (range <= 10) { return 1; }
    if (range <= 100) { return 5; }
    if (range <= 1000) { return 10; }
    if (range <= 10000) { return 100; }
    if (range <= 1000000) { return 1000; }
    return 10000;
  }, [config, rangeMin, rangeMax]);

  // Get current range from active filters
  const currentActive: IActiveFilter | undefined = React.useMemo(function (): IActiveFilter | undefined {
    for (let i: number = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        return activeFilters[i];
      }
    }
    return undefined;
  }, [activeFilters, filterName]);

  const parsedRange = currentActive ? parseRangeToken(currentActive.value) : undefined;
  const initialStart: number = parsedRange && isFinite(parsedRange.min as number) ? (parsedRange.min as number) : rangeMin;
  const initialEnd: number = parsedRange && isFinite(parsedRange.max as number) ? (parsedRange.max as number) : rangeMax;

  const [currentRange, setCurrentRange] = React.useState<[number, number]>([initialStart, initialEnd]);

  React.useEffect(function (): void {
    setCurrentRange([initialStart, initialEnd]);
  }, [initialStart, initialEnd]);

  function handleValueChanged(e: { value?: number[] }): void {
    if (!Array.isArray(e.value) || e.value.length < 2) {
      return;
    }
    const nextRange: [number, number] = [e.value[0], e.value[1]];
    setCurrentRange(nextRange);

    // If reset to full range, remove filter
    if (nextRange[0] === rangeMin && nextRange[1] === rangeMax) {
      if (currentActive) {
        onToggleRefiner(currentActive);
      }
      return;
    }

    // Build and dispatch FQL range token
    const fqlToken: string = 'range(decimal(' + nextRange[0] + '), decimal(' + nextRange[1] + '))';
    onToggleRefiner({
      filterName: filterName,
      value: fqlToken,
      operator: 'AND'
    });
  }

  const formatConfig: IFilterConfig = config || {
    id: filterName,
    displayName: filterName,
    managedProperty: filterName,
    filterType: 'slider',
    operator: 'AND',
    maxValues: 1,
    defaultExpanded: true,
    showCount: false,
    sortBy: 'count',
    sortDirection: 'desc',
  };
  if (!formatConfig.rangeFormat && filterName.toLowerCase().indexOf('size') >= 0) {
    formatConfig.rangeFormat = 'bytes';
  }

  return (
    <div
      className={styles.sliderFilterContainer}
      role="group"
      aria-label={(config ? config.displayName : filterName) + ' range filter'}
    >
      <div className={styles.sliderLabels}>
        <span className={styles.sliderMinLabel}>{formatNumericValue(currentRange[0], formatConfig)}</span>
        <span className={styles.sliderMaxLabel}>{formatNumericValue(currentRange[1], formatConfig)}</span>
      </div>
      <RangeSlider
        min={rangeMin}
        max={rangeMax}
        step={step}
        value={currentRange}
        onValueChanged={handleValueChanged}
        showRange={true}
      >
        <Label visible={true} />
        <Tooltip enabled={true} showMode="always" />
      </RangeSlider>
    </div>
  );
};

export default SliderFilter;
