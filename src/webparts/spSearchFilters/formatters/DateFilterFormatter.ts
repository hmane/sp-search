import type { IFilterValueFormatter, IFilterConfig } from '@interfaces/index';

/**
 * Parse an FQL range token for dates.
 * Expected format: range(datetime("2026-01-01T00:00:00Z"), datetime("2026-12-31T23:59:59Z"))
 */
function parseDateRange(token: string): { from: Date; to: Date } | undefined {
  const regex: RegExp = /^range\(datetime\("([^"]+)"\),\s*datetime\("([^"]+)"\)\)$/;
  const match: RegExpMatchArray | null = token.match(regex);
  if (!match) {
    return undefined;
  }
  const from: Date = new Date(match[1]);
  const to: Date = new Date(match[2]);
  if (isNaN(from.getTime()) || isNaN(to.getTime())) {
    return undefined;
  }
  return { from: from, to: to };
}

/**
 * Format a Date for short display. Format: "MMM D, YYYY"
 */
function formatDate(date: Date): string {
  const months: string[] = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
  ];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

/**
 * DateFilterFormatter — formats date range refinement values.
 *
 * Handles FQL range(datetime(...), datetime(...)) tokens.
 * Displays as "Jan 1, 2026 – Dec 31, 2026" in the pill bar.
 */
export const DateFilterFormatter: IFilterValueFormatter = {
  id: 'daterange',

  formatForDisplay: function (rawValue: string, _config: IFilterConfig): string {
    const parsed = parseDateRange(rawValue);
    if (!parsed) {
      return rawValue;
    }
    return formatDate(parsed.from) + ' – ' + formatDate(parsed.to);
  },

  formatForQuery: function (displayValue: unknown, _config: IFilterConfig): string {
    // Date ranges are already in FQL format
    return String(displayValue);
  },

  formatForUrl: function (rawValue: string): string {
    return encodeURIComponent(rawValue);
  },

  parseFromUrl: function (urlValue: string): string {
    return decodeURIComponent(urlValue);
  }
};
