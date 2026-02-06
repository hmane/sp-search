import * as React from 'react';
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

type DatePreset = 'today' | 'thisWeek' | 'thisMonth' | 'thisQuarter' | 'thisYear' | 'custom';

interface IDateRange {
  from: Date;
  to: Date;
}

/** Preset labels for display. */
const PRESET_LABELS: Record<DatePreset, string> = {
  today: 'Today',
  thisWeek: 'This Week',
  thisMonth: 'This Month',
  thisQuarter: 'This Quarter',
  thisYear: 'This Year',
  custom: 'Custom'
};

/** All preset keys in display order. */
const PRESET_KEYS: DatePreset[] = [
  'today',
  'thisWeek',
  'thisMonth',
  'thisQuarter',
  'thisYear',
  'custom'
];

/**
 * Formats a Date to an ISO 8601 UTC string suitable for FQL range() tokens.
 * Output: "2026-01-15T00:00:00Z"
 */
function toFqlDateString(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * Builds an FQL range() token for SharePoint Search date refinement.
 * Format: range(datetime("2026-01-01T00:00:00Z"), datetime("2026-12-31T23:59:59Z"))
 */
function buildFqlRange(from: Date, to: Date): string {
  return 'range(datetime("' + toFqlDateString(from) + '"), datetime("' + toFqlDateString(to) + '"))';
}

/**
 * Computes the UTC start-of-day and end-of-day for a given date.
 * Uses the local date components to derive UTC boundaries.
 */
function getStartOfDayUTC(date: Date): Date {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0));
}

function getEndOfDayUTC(date: Date): Date {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 0));
}

/**
 * Calculates the date range for a given preset relative to today.
 */
function calculatePresetRange(preset: DatePreset): IDateRange | undefined {
  const now: Date = new Date();
  const year: number = now.getFullYear();
  const month: number = now.getMonth();
  const day: number = now.getDate();
  const dayOfWeek: number = now.getDay(); // 0=Sun, 1=Mon, ...

  switch (preset) {
    case 'today': {
      return {
        from: getStartOfDayUTC(now),
        to: getEndOfDayUTC(now)
      };
    }
    case 'thisWeek': {
      // Monday of current week (ISO week starts on Monday)
      const diffToMonday: number = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
      const monday: Date = new Date(year, month, day - diffToMonday);
      const sunday: Date = new Date(year, month, day - diffToMonday + 6);
      return {
        from: getStartOfDayUTC(monday),
        to: getEndOfDayUTC(sunday)
      };
    }
    case 'thisMonth': {
      const firstOfMonth: Date = new Date(year, month, 1);
      const lastOfMonth: Date = new Date(year, month + 1, 0); // Day 0 of next month = last day of this month
      return {
        from: getStartOfDayUTC(firstOfMonth),
        to: getEndOfDayUTC(lastOfMonth)
      };
    }
    case 'thisQuarter': {
      // Q1=Jan-Mar, Q2=Apr-Jun, Q3=Jul-Sep, Q4=Oct-Dec
      const quarterStartMonth: number = Math.floor(month / 3) * 3;
      const firstOfQuarter: Date = new Date(year, quarterStartMonth, 1);
      const lastOfQuarter: Date = new Date(year, quarterStartMonth + 3, 0);
      return {
        from: getStartOfDayUTC(firstOfQuarter),
        to: getEndOfDayUTC(lastOfQuarter)
      };
    }
    case 'thisYear': {
      const firstOfYear: Date = new Date(year, 0, 1);
      const lastOfYear: Date = new Date(year, 11, 31);
      return {
        from: getStartOfDayUTC(firstOfYear),
        to: getEndOfDayUTC(lastOfYear)
      };
    }
    case 'custom': {
      return undefined;
    }
    default: {
      return undefined;
    }
  }
}

/**
 * Attempts to parse an FQL range token back into from/to dates.
 * Expected format: range(datetime("..."), datetime("..."))
 * Returns undefined if parsing fails.
 */
function parseFqlRange(fqlToken: string): IDateRange | undefined {
  // Match: range(datetime("..."), datetime("..."))
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
 * Returns 'custom' if no preset matches.
 */
function detectPresetFromFql(fqlToken: string): DatePreset {
  const parsed: IDateRange | undefined = parseFqlRange(fqlToken);
  if (!parsed) {
    return 'custom';
  }

  const presetKeys: DatePreset[] = ['today', 'thisWeek', 'thisMonth', 'thisQuarter', 'thisYear'];
  for (let i: number = 0; i < presetKeys.length; i++) {
    const presetRange: IDateRange | undefined = calculatePresetRange(presetKeys[i]);
    if (presetRange) {
      if (
        parsed.from.getTime() === presetRange.from.getTime() &&
        parsed.to.getTime() === presetRange.to.getTime()
      ) {
        return presetKeys[i];
      }
    }
  }

  return 'custom';
}

/**
 * Formats a Date for display in local timezone. Format: "MMM D"
 */
function formatDateForDisplay(date: Date): string {
  const months: string[] = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
  ];
  return months[date.getMonth()] + ' ' + date.getDate();
}

/**
 * Formats a Date as YYYY-MM-DD for use in native HTML date inputs.
 * Uses local date components.
 */
function toInputDateString(date: Date): string {
  const y: string = String(date.getFullYear());
  const monthNum: number = date.getMonth() + 1;
  const dayNum: number = date.getDate();
  const m: string = monthNum < 10 ? '0' + String(monthNum) : String(monthNum);
  const d: string = dayNum < 10 ? '0' + String(dayNum) : String(dayNum);
  return y + '-' + m + '-' + d;
}

/**
 * Parses a YYYY-MM-DD input string into a Date (local time).
 * Returns undefined for invalid values.
 */
function parseInputDateString(value: string): Date | undefined {
  if (!value) {
    return undefined;
  }
  const parts: string[] = value.split('-');
  if (parts.length !== 3) {
    return undefined;
  }
  const year: number = parseInt(parts[0], 10);
  const month: number = parseInt(parts[1], 10) - 1;
  const day: number = parseInt(parts[2], 10);
  if (isNaN(year) || isNaN(month) || isNaN(day)) {
    return undefined;
  }
  return new Date(year, month, day);
}

/**
 * DateRangeFilter renders preset date range buttons and a custom date picker.
 * It emits FQL range() tokens via onToggleRefiner for SharePoint Search refinement.
 */
const DateRangeFilter: React.FC<IDateRangeFilterProps> = (props: IDateRangeFilterProps): React.ReactElement => {
  const { filterName, activeFilters, onToggleRefiner } = props;

  // Find the currently active date filter for this filterName (only one date range active at a time)
  const currentActiveFilter: IActiveFilter | undefined = React.useMemo(function (): IActiveFilter | undefined {
    for (let i: number = 0; i < activeFilters.length; i++) {
      if (activeFilters[i].filterName === filterName) {
        return activeFilters[i];
      }
    }
    return undefined;
  }, [activeFilters, filterName]);

  // Detect the active preset from the current FQL token
  const activePreset: DatePreset | undefined = React.useMemo(function (): DatePreset | undefined {
    if (!currentActiveFilter) {
      return undefined;
    }
    return detectPresetFromFql(currentActiveFilter.value);
  }, [currentActiveFilter]);

  // Custom date range state — initialized from active filter if it's a custom range
  const initialCustomRange: IDateRange | undefined = React.useMemo(function (): IDateRange | undefined {
    if (currentActiveFilter && activePreset === 'custom') {
      return parseFqlRange(currentActiveFilter.value);
    }
    return undefined;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const [customFrom, setCustomFrom] = React.useState<string>(
    initialCustomRange ? toInputDateString(initialCustomRange.from) : ''
  );
  const [customTo, setCustomTo] = React.useState<string>(
    initialCustomRange ? toInputDateString(initialCustomRange.to) : ''
  );
  const [showCustomInputs, setShowCustomInputs] = React.useState<boolean>(activePreset === 'custom');

  /**
   * Dispatches a date range filter. Since date refiners are mutually exclusive
   * (only one range at a time), we first remove any existing filter for this
   * filterName and then add the new one. Toggling the same preset again removes it.
   */
  function dispatchDateFilter(fqlToken: string): void {
    const filter: IActiveFilter = {
      filterName: filterName,
      value: fqlToken,
      operator: 'AND'
    };
    onToggleRefiner(filter);
  }

  /**
   * Handles clicking a preset button. If the same preset is already active,
   * clicking it again deselects (toggles off).
   */
  function handlePresetClick(preset: DatePreset): void {
    if (preset === 'custom') {
      setShowCustomInputs(true);
      // If we already have custom dates, apply them
      if (customFrom && customTo) {
        applyCustomRange(customFrom, customTo);
      }
      return;
    }

    setShowCustomInputs(false);

    // If this preset is already active, toggle it off
    if (activePreset === preset && currentActiveFilter) {
      dispatchDateFilter(currentActiveFilter.value);
      return;
    }

    const range: IDateRange | undefined = calculatePresetRange(preset);
    if (range) {
      const fqlToken: string = buildFqlRange(range.from, range.to);
      dispatchDateFilter(fqlToken);
    }
  }

  /**
   * Builds and dispatches a custom date range from the from/to input values.
   */
  function applyCustomRange(fromValue: string, toValue: string): void {
    const fromDate: Date | undefined = parseInputDateString(fromValue);
    const toDate: Date | undefined = parseInputDateString(toValue);

    if (!fromDate || !toDate) {
      return;
    }

    // Ensure from <= to
    if (fromDate.getTime() > toDate.getTime()) {
      return;
    }

    const fromUTC: Date = getStartOfDayUTC(fromDate);
    const toUTC: Date = getEndOfDayUTC(toDate);
    const fqlToken: string = buildFqlRange(fromUTC, toUTC);
    dispatchDateFilter(fqlToken);
  }

  function handleFromChange(ev: React.ChangeEvent<HTMLInputElement>): void {
    const value: string = ev.target.value;
    setCustomFrom(value);
    if (value && customTo) {
      applyCustomRange(value, customTo);
    }
  }

  function handleToChange(ev: React.ChangeEvent<HTMLInputElement>): void {
    const value: string = ev.target.value;
    setCustomTo(value);
    if (customFrom && value) {
      applyCustomRange(customFrom, value);
    }
  }

  /** Build the display string for a custom active filter (for the active state label). */
  function getCustomDisplayLabel(): string {
    if (!currentActiveFilter) {
      return '';
    }
    const parsed: IDateRange | undefined = parseFqlRange(currentActiveFilter.value);
    if (!parsed) {
      return '';
    }
    return 'Custom: ' + formatDateForDisplay(parsed.from) + ' - ' + formatDateForDisplay(parsed.to);
  }

  return (
    <div className={styles.dateRangeContainer} role="group" aria-label={filterName + ' date range filter'}>
      {/* Preset buttons row */}
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
              aria-label={PRESET_LABELS[preset] + ' date range'}
            >
              {PRESET_LABELS[preset]}
            </button>
          );
        })}
      </div>

      {/* Custom date range picker */}
      {showCustomInputs && (
        <div className={styles.customRange}>
          <div className={styles.dateInputGroup}>
            <label className={styles.dateInputLabel} htmlFor={filterName + '-from'}>
              From
            </label>
            <input
              id={filterName + '-from'}
              type="date"
              className={styles.dateInput}
              value={customFrom}
              onChange={handleFromChange}
              aria-label="Start date"
              max={customTo || undefined}
            />
          </div>
          <div className={styles.dateInputGroup}>
            <label className={styles.dateInputLabel} htmlFor={filterName + '-to'}>
              To
            </label>
            <input
              id={filterName + '-to'}
              type="date"
              className={styles.dateInput}
              value={customTo}
              onChange={handleToChange}
              aria-label="End date"
              min={customFrom || undefined}
            />
          </div>
        </div>
      )}

      {/* Active state display */}
      {currentActiveFilter && activePreset === 'custom' && (
        <div className={styles.dateRangeActiveLabel}>
          {getCustomDisplayLabel()}
        </div>
      )}
    </div>
  );
};

export default DateRangeFilter;
