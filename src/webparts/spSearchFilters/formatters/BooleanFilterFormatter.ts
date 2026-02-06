import type { IFilterValueFormatter, IFilterConfig } from '@interfaces/index';

/**
 * BooleanFilterFormatter — formats boolean refinement values.
 * Maps "1" → "Yes", "0" → "No" for display.
 * FQL tokens use quoted string values: '"1"' and '"0"'.
 */
export const BooleanFilterFormatter: IFilterValueFormatter = {
  id: 'toggle',

  formatForDisplay: function (rawValue: string, _config: IFilterConfig): string {
    const cleaned: string = rawValue.replace(/"/g, '');
    if (cleaned === '1' || cleaned.toLowerCase() === 'true') {
      return 'Yes';
    }
    if (cleaned === '0' || cleaned.toLowerCase() === 'false') {
      return 'No';
    }
    return rawValue;
  },

  formatForQuery: function (displayValue: unknown, _config: IFilterConfig): string {
    const val: string = String(displayValue).replace(/"/g, '');
    if (val === '1' || val.toLowerCase() === 'true' || val.toLowerCase() === 'yes') {
      return '"1"';
    }
    if (val === '0' || val.toLowerCase() === 'false' || val.toLowerCase() === 'no') {
      return '"0"';
    }
    return '"' + val + '"';
  },

  formatForUrl: function (rawValue: string): string {
    return rawValue.replace(/"/g, '');
  },

  parseFromUrl: function (urlValue: string): string {
    return '"' + urlValue + '"';
  }
};
