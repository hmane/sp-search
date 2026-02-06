import type { IFilterValueFormatter, IFilterConfig } from '@interfaces/index';

/**
 * DefaultFilterFormatter — pass-through formatter for checkbox and other
 * simple string-value filters. Values are used as-is with no transformation.
 */
export const DefaultFilterFormatter: IFilterValueFormatter = {
  id: 'checkbox',

  formatForDisplay: function (rawValue: string, _config: IFilterConfig): string {
    // Remove surrounding quotes if present
    if (rawValue.length >= 2 && rawValue.charAt(0) === '"' && rawValue.charAt(rawValue.length - 1) === '"') {
      return rawValue.substring(1, rawValue.length - 1);
    }
    return rawValue;
  },

  formatForQuery: function (displayValue: unknown, _config: IFilterConfig): string {
    const val: string = String(displayValue);
    // Wrap in quotes if not already quoted
    if (val.charAt(0) !== '"') {
      return '"' + val + '"';
    }
    return val;
  },

  formatForUrl: function (rawValue: string): string {
    return encodeURIComponent(rawValue);
  },

  parseFromUrl: function (urlValue: string): string {
    return decodeURIComponent(urlValue);
  }
};
