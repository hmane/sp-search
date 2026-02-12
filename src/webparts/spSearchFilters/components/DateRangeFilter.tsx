import * as React from 'react';
import { DateBox } from 'devextreme-react/date-box';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SpSearchFilters.module.scss';
import type {
  IRefinerValue,
  IActiveFilter,
  IFilterConfig
} from '@interfaces/index';

export interface IDateRangeFilterProps {
  filterName: string;
  values: IRefinerValue[];
  config: IFilterConfig | undefined;
  activeFilters: IActiveFilter[];
  onToggleRefiner: (filter: IActiveFilter) => void;
}

type DatePreset = 'today' | 'last7' | 'last30' | 'thisMonth' | 'last3months' | 'thisYear' | 'custom';

interface IDateRange {
  from: Date;
  to: Date;
}

/** Preset labels — compact for pill display. */
const PRESET_LABELS: Record<DatePreset, string> = {
  today: 'Today',
  last7: '7 days',
  last30: '30 days',
  thisMonth: 'This month',
  last3months: '3 months',
  thisYear: 'This year',
  custom: 'Custom'
};

/** Full display labels for active filter indicator. */
const PRESET_DISPLAY: Record<DatePreset, string> = {
  today: 'Today',
  last7: 'Last 7 days',
  last30: 'Last 30 days',
  thisMonth: 'This month',
  last3months: 'Last 3 months',
  thisYear: 'This year',
  custom: 'Custom range'
};

/** All preset keys in display order. */
const PRESET_KEYS: DatePreset[] = [
  'today', 'last7', 'last30', 'thisMonth', 'last3months', 'thisYear', 'custom'
];

const MONTHS: string[] = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

// ── Date utility functions ───────────────────────────────────────────────

/**
 * Formats a Date to ISO 8601 UTC string suitable for FQL range() tokens.
 * Output: "2026-01-15T00:00:00Z"
 */
function toFqlDateString(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * Builds an FQL range() token for SharePoint Search date refinement.
 */
function buildFqlRange(from: Date, to: Date): string {
  return 'range(datetime("' + toFqlDateString(from) + '"), datetime("' + toFqlDateString(to) + '"))';
}

function getStartOfDayUTC(date: Date): Date {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0));
}

function getEndOfDayUTC(date: Date): Date {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 0));
}

/** Short date display: "Jan 15" */
function formatDateShort(date: Date): string {
  return MONTHS[date.getMonth()] + ' ' + date.getDate();
}

/** Full date display: "Jan 15, 2026" */
function formatDateFull(date: Date): string {
  return MONTHS[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

// ── Preset range calculation ─────────────────────────────────────────────

function calculatePresetRange(preset: DatePreset): IDateRange | undefined {
  const now: Date = new Date();
  const year: number = now.getFullYear();
  const month: number = now.getMonth();
  const day: number = now.getDate();

  switch (preset) {
    case 'today':
      return { from: getStartOfDayUTC(now), to: getEndOfDayUTC(now) };
    case 'last7': {
      const from: Date = new Date(year, month, day - 6);
      return { from: getStartOfDayUTC(from), to: getEndOfDayUTC(now) };
    }
    case 'last30': {
      const from: Date = new Date(year, month, day - 29);
      return { from: getStartOfDayUTC(from), to: getEndOfDayUTC(now) };
    }
    case 'thisMonth': {
      return {
        from: getStartOfDayUTC(new Date(year, month, 1)),
        to: getEndOfDayUTC(new Date(year, month + 1, 0))
      };
    }
    case 'last3months': {
      return {
        from: getStartOfDayUTC(new Date(year, month - 2, 1)),
        to: getEndOfDayUTC(now)
      };
    }
    case 'thisYear': {
      return {
        from: getStartOfDayUTC(new Date(year, 0, 1)),
        to: getEndOfDayUTC(new Date(year, 11, 31))
      };
    }
    case 'custom':
      return undefined;
    default:
      return undefined;
  }
}

// ── FQL parsing ──────────────────────────────────────────────────────────

/**
 * Attempts to parse an FQL range token back into from/to dates.
 */
function parseFqlRange(fqlToken: string): IDateRange | undefined {
  const regex: RegExp = /^range\(datetime\("([^"]+)"\),\s*datetime\("([^"]+)"\)\)$/;
  const match: RegExpMatchArray | null = fqlToken.match(regex);
  if (!match) {
    return undefined;
  }
  const fromDate: Date = new Date(match[1]);
  const toDate: Date = new Date(match[2]);
  if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
    return undefined;
  }
  return { from: fromDate, to: toDate };
}

/**
 * Determines which preset (if any) matches the given FQL range token.
 */
function detectPresetFromFql(fqlToken: string): DatePreset {
  const parsed: IDateRange | undefined = parseFqlRange(fqlToken);
  if (!parsed) {
    return 'custom';
  }
  const keys: DatePreset[] = ['today', 'last7', 'last30', 'thisMonth', 'last3months', 'thisYear'];
  for (let i: number = 0; i < keys.length; i++) {
    const range: IDateRange | undefined = calculatePresetRange(keys[i]);
    if (range && parsed.from.getTime() === range.from.getTime() && parsed.to.getTime() === range.to.getTime()) {
      return keys[i];
    }
  }
  return 'custom';
}

// ── Component ────────────────────────────────────────────────────────────

const DateRangeFilter: React.FC<IDateRangeFilterProps> = (props: IDateRangeFilterProps): React.ReactElement => {
  const { filterName, activeFilters, onToggleRefiner } = props;

  // Find the currently active date filter (only one at a time)
  const currentActiveFilter: IActiveFilter | undefined = React.useMemo(function (): IActiveFilter | undefined {
    for (let i: number = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        return activeFilters[i];
      }
    }
    return undefined;
  }, [activeFilters, filterName]);

  // Detect active preset from current FQL token
  const activePreset: DatePreset | undefined = React.useMemo(function (): DatePreset | undefined {
    if (!currentActiveFilter) {
      return undefined;
    }
    return detectPresetFromFql(currentActiveFilter.value);
  }, [currentActiveFilter]);

  // Custom date range state
  const initialCustomRange: IDateRange | undefined = React.useMemo(function (): IDateRange | undefined {
    if (currentActiveFilter && activePreset === 'custom') {
      return parseFqlRange(currentActiveFilter.value);
    }
    return undefined;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const [customFrom, setCustomFrom] = React.useState<Date | undefined>(
    initialCustomRange ? initialCustomRange.from : undefined
  );
  const [customTo, setCustomTo] = React.useState<Date | undefined>(
    initialCustomRange ? initialCustomRange.to : undefined
  );
  const [showCustomInputs, setShowCustomInputs] = React.useState<boolean>(activePreset === 'custom');

  // ── Handlers ─────────────────────────────────────────────

  function dispatchDateFilter(fqlToken: string, displayValue?: string): void {
    onToggleRefiner({
      filterName: filterName,
      value: fqlToken,
      displayValue: displayValue || undefined,
      operator: 'AND'
    });
  }

  function handlePresetClick(preset: DatePreset): void {
    if (preset === 'custom') {
      setShowCustomInputs(true);
      if (customFrom && customTo) {
        applyCustomRangeDates(customFrom, customTo);
      }
      return;
    }

    setShowCustomInputs(false);

    // Toggle off if same preset is already active
    if (activePreset === preset && currentActiveFilter) {
      dispatchDateFilter(currentActiveFilter.value);
      return;
    }

    const range: IDateRange | undefined = calculatePresetRange(preset);
    if (range) {
      dispatchDateFilter(buildFqlRange(range.from, range.to), PRESET_DISPLAY[preset]);
    }
  }

  function handleClearFilter(): void {
    if (currentActiveFilter) {
      dispatchDateFilter(currentActiveFilter.value);
    }
    setShowCustomInputs(false);
    setCustomFrom(undefined);
    setCustomTo(undefined);
  }

  function applyCustomRangeDates(fromDate: Date, toDate: Date): void {
    if (fromDate.getTime() > toDate.getTime()) {
      return;
    }
    const fromUTC: Date = getStartOfDayUTC(fromDate);
    const toUTC: Date = getEndOfDayUTC(toDate);
    dispatchDateFilter(
      buildFqlRange(fromUTC, toUTC),
      formatDateShort(fromDate) + ' \u2013 ' + formatDateShort(toDate)
    );
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function handleFromChanged(e: any): void {
    const value: Date | undefined = e.value || undefined;
    setCustomFrom(value);
    if (value && customTo) {
      applyCustomRangeDates(value, customTo);
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function handleToChanged(e: any): void {
    const value: Date | undefined = e.value || undefined;
    setCustomTo(value);
    if (customFrom && value) {
      applyCustomRangeDates(customFrom, value);
    }
  }

  /** Derive the display text for the active filter indicator. */
  function getActiveDisplayText(): string {
    if (!currentActiveFilter) {
      return '';
    }
    if (currentActiveFilter.displayValue) {
      return currentActiveFilter.displayValue;
    }
    if (activePreset && activePreset !== 'custom') {
      return PRESET_DISPLAY[activePreset];
    }
    const parsed: IDateRange | undefined = parseFqlRange(currentActiveFilter.value);
    if (parsed) {
      return formatDateFull(parsed.from) + ' \u2013 ' + formatDateFull(parsed.to);
    }
    return 'Custom range';
  }

  // ── Render ───────────────────────────────────────────────

  return (
    <div className={styles.dateRangeContainer} role="group" aria-label={filterName + ' date range filter'}>

      {/* Active filter indicator with clear button */}
      {currentActiveFilter && (
        <div className={styles.dateActiveBar}>
          <Icon iconName="Clock" className={styles.dateActiveIcon} />
          <span className={styles.dateActiveText} title={getActiveDisplayText()}>
            {getActiveDisplayText()}
          </span>
          <button
            type="button"
            className={styles.dateActiveClear}
            onClick={handleClearFilter}
            aria-label="Clear date filter"
            title="Clear"
          >
            <Icon iconName="Cancel" />
          </button>
        </div>
      )}

      {/* Preset pill buttons */}
      <div className={styles.presetButtons}>
        {PRESET_KEYS.map(function (preset: DatePreset): React.ReactElement {
          const isActive: boolean = preset === 'custom'
            ? (showCustomInputs && activePreset === 'custom')
            : activePreset === preset;

          return (
            <button
              key={preset}
              type="button"
              className={isActive ? styles.presetButtonActive : styles.presetButton}
              onClick={function (): void { handlePresetClick(preset); }}
              aria-pressed={isActive}
              aria-label={PRESET_DISPLAY[preset]}
            >
              {PRESET_LABELS[preset]}
            </button>
          );
        })}
      </div>

      {/* Custom date range picker */}
      {showCustomInputs && (
        <div className={styles.customRangeSection}>
          <div className={styles.customRangeRow}>
            <label className={styles.customRangeLabel}>From</label>
            <div className={styles.dateInputGroup}>
              <DateBox
                value={customFrom}
                onValueChanged={handleFromChanged}
                type="date"
                displayFormat="MMM d, yyyy"
                showClearButton={true}
                max={customTo}
                placeholder="Start date"
                width="100%"
              />
            </div>
          </div>
          <div className={styles.customRangeRow}>
            <label className={styles.customRangeLabel}>To</label>
            <div className={styles.dateInputGroup}>
              <DateBox
                value={customTo}
                onValueChanged={handleToChanged}
                type="date"
                displayFormat="MMM d, yyyy"
                showClearButton={true}
                min={customFrom}
                placeholder="End date"
                width="100%"
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default DateRangeFilter;
